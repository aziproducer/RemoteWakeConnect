using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using RemoteWakeConnect.Models;
using RemoteWakeConnect.Native;

namespace RemoteWakeConnect.Services
{
    /// <summary>
    /// ホストごとのフラグ管理
    /// </summary>
    public class HostFlags
    {
        /// <summary>
        /// WTSEnumerateSessionsExがそのホストで使えるか（null=未判定）
        /// </summary>
        public bool? SupportsWtsEx { get; set; }

        /// <summary>
        /// WTSEnumerateをしばらく打たない期限（0件や到達不可などの負キャッシュ）
        /// </summary>
        public DateTime NegativeUntil { get; set; } = DateTime.MinValue;
        
        /// <summary>
        /// 連続失敗回数（リトライ間隔の調整用）
        /// </summary>
        public int ConsecutiveFailures { get; set; } = 0;
    }

    /// <summary>
    /// セッション監視サービス
    /// </summary>
    public class SessionMonitorService
    {
        private readonly DebugLogService _debugLog;
        private readonly ConcurrentDictionary<string, IntPtr> _serverHandles = new ConcurrentDictionary<string, IntPtr>();
        private readonly ConcurrentDictionary<string, DateTime> _serverHandleLastUsed = new ConcurrentDictionary<string, DateTime>();
        private readonly ConcurrentDictionary<string, OsInfo> _osInfoCache = new ConcurrentDictionary<string, OsInfo>();
        private readonly ConcurrentDictionary<string, DateTime> _osInfoCacheTime = new ConcurrentDictionary<string, DateTime>();
        private readonly TimeSpan _osInfoCacheTTL = TimeSpan.FromMinutes(10);
        
        // ハンドルプーリング設定
        private readonly TimeSpan _handleKeepAlive = TimeSpan.FromMinutes(30); // 30分間保持
        private readonly TimeSpan _handleCleanupInterval = TimeSpan.FromMinutes(5); // 5分ごとにクリーンアップ
        private DateTime _lastHandleCleanup = DateTime.Now;
        
        // 新規: ホストごとのEx可否と負キャッシュ
        private readonly ConcurrentDictionary<string, HostFlags> _hostFlags = new ConcurrentDictionary<string, HostFlags>();
        
        // 削除：個別のNegativeTtlは使わず、失敗回数に応じて動的に決定

        // FQDN解決キャッシュ
        private readonly ConcurrentDictionary<string, string> _fqdnCache = new ConcurrentDictionary<string, string>();
        private readonly TimeSpan _fqdnCacheTTL = TimeSpan.FromHours(1);
        private readonly ConcurrentDictionary<string, DateTime> _fqdnCacheTime = new ConcurrentDictionary<string, DateTime>();
        
        public SessionMonitorService()
        {
            _debugLog = new DebugLogService();
        }

        /// <summary>
        /// セッション情報を確認
        /// </summary>
        /// <param name="hostName">ホスト名またはIPアドレス</param>
        /// <param name="customPort">カスタムポート番号（デフォルト3389）</param>
        /// <param name="cachedOsInfo">キャッシュされたOS情報（ある場合）</param>
        /// <returns>セッション確認結果</returns>
        public async Task<SessionCheckResult> CheckSessionsAsync(string hostName, int customPort = 3389, OsInfo? cachedOsInfo = null)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            _debugLog.WriteLine($"[SessionMonitor] CheckSessionsAsync START for: {hostName}");
            
            var result = new SessionCheckResult
            {
                CurrentUser = Environment.UserName
            };

            try
            {
                // ポート生存確認（200msタイムアウト）- カスタムポートとRPC(135)の両方を確認
                _debugLog.WriteLine($"[SessionMonitor] Checking port availability (RDP:{customPort}, RPC:135)... (elapsed: {sw.ElapsedMilliseconds}ms)");
                
                // RDPポート(カスタム)とRPCポート(135)を並列チェック
                var rdpCheckTask = CheckPortAsync(hostName, customPort, 200);
                var rpcCheckTask = CheckPortAsync(hostName, 135, 200);
                await Task.WhenAll(rdpCheckTask, rpcCheckTask);
                
                bool isRdpReachable = rdpCheckTask.Result;
                bool isRpcReachable = rpcCheckTask.Result;
                
                _debugLog.WriteLine($"[SessionMonitor] Port check results - RDP({customPort}): {isRdpReachable}, RPC(135): {isRpcReachable} (elapsed: {sw.ElapsedMilliseconds}ms)");
                
                if (!isRdpReachable && !isRpcReachable)
                {
                    _debugLog.WriteLine($"[SessionMonitor] Port {customPort} not reachable, skipping (elapsed: {sw.ElapsedMilliseconds}ms)");
                    result.OsInfo = GetDefaultOsInfo();
                    result.Sessions = Array.Empty<SessionInfo>();
                    result.ErrorMessage = "接続先に到達できません";
                    
                    // ポートに到達できない場合は負のキャッシュを設定
                    var flags = _hostFlags.GetOrAdd(hostName, _ => new HostFlags());
                    flags.ConsecutiveFailures++;
                    
                    // 失敗回数に応じてリトライ間隔を調整
                    TimeSpan retryInterval;
                    if (flags.ConsecutiveFailures == 1)
                    {
                        retryInterval = TimeSpan.FromSeconds(5);  // 1回目の失敗後: 5秒
                    }
                    else if (flags.ConsecutiveFailures == 2)
                    {
                        retryInterval = TimeSpan.FromSeconds(10); // 2回目の失敗後: 10秒
                    }
                    else
                    {
                        retryInterval = TimeSpan.FromSeconds(20); // 3回目以降: 20秒
                    }
                    
                    flags.NegativeUntil = DateTime.Now.Add(retryInterval);
                    _debugLog.WriteLine($"[SessionMonitor] Port unreachable (failure #{flags.ConsecutiveFailures}), retry after {retryInterval.TotalSeconds}s at {flags.NegativeUntil:HH:mm:ss}");
                    
                    return result;
                }
                
                // ポートに到達可能な場合は負のキャッシュと失敗回数をクリア
                if (_hostFlags.TryGetValue(hostName, out var hostFlags))
                {
                    if (hostFlags.NegativeUntil > DateTime.MinValue || hostFlags.ConsecutiveFailures > 0)
                    {
                        _debugLog.WriteLine($"[SessionMonitor] Port is reachable, clearing negative cache and resetting failure count (was {hostFlags.ConsecutiveFailures})");
                        hostFlags.NegativeUntil = DateTime.MinValue;
                        hostFlags.ConsecutiveFailures = 0;
                    }
                }
                
                // OS情報を取得（YAMLキャッシュ、メモリキャッシュ、デフォルトの順）
                if (cachedOsInfo != null)
                {
                    // YAMLから読み込んだOS情報を使用
                    result.OsInfo = cachedOsInfo;
                    _osInfoCache.AddOrUpdate(hostName, cachedOsInfo, (k, v) => cachedOsInfo);
                    _osInfoCacheTime.AddOrUpdate(hostName, DateTime.Now, (k, v) => DateTime.Now);
                    _debugLog.WriteLine($"[SessionMonitor] Using YAML cached OS info: {cachedOsInfo.Type} (elapsed: {sw.ElapsedMilliseconds}ms)");
                }
                else if (_osInfoCache.TryGetValue(hostName, out var memCachedInfo))
                {
                    // メモリキャッシュから取得
                    result.OsInfo = memCachedInfo;
                    _debugLog.WriteLine($"[SessionMonitor] Using memory cached OS info: {memCachedInfo.Type} (elapsed: {sw.ElapsedMilliseconds}ms)");
                }
                else
                {
                    // デフォルト値を使用
                    result.OsInfo = GetDefaultOsInfo();
                    _osInfoCache.AddOrUpdate(hostName, result.OsInfo, (k, v) => result.OsInfo);
                    _osInfoCacheTime.AddOrUpdate(hostName, DateTime.Now, (k, v) => DateTime.Now);
                    _debugLog.WriteLine($"[SessionMonitor] Using default OS info (elapsed: {sw.ElapsedMilliseconds}ms)");
                }

                // セッション情報を取得
                _debugLog.WriteLine($"[SessionMonitor] Getting sessions... (elapsed: {sw.ElapsedMilliseconds}ms)");
                result.Sessions = await GetSessionsAsync(hostName);
                _debugLog.WriteLine($"[SessionMonitor] Found {result.Sessions.Length} sessions (elapsed: {sw.ElapsedMilliseconds}ms)");

                // 他のユーザーが使用中か判定
                result.IsInUseByOthers = IsInUseByOthers(result.Sessions, result.CurrentUser);
                _debugLog.WriteLine($"[SessionMonitor] IsInUseByOthers: {result.IsInUseByOthers} (elapsed: {sw.ElapsedMilliseconds}ms)");

                // 警告メッセージを生成
                result.WarningMessage = GenerateWarningMessage(result);

                // 接続可能か判定
                result.CanConnect = DetermineConnectability(result);
                
                _debugLog.WriteLine($"[SessionMonitor] CheckSessionsAsync COMPLETED successfully (total elapsed: {sw.ElapsedMilliseconds}ms)");
                return result;
            }
            catch (Exception ex)
            {
                _debugLog.WriteLine($"[SessionMonitor] Error: {ex.Message} (elapsed: {sw.ElapsedMilliseconds}ms)");
                result.ErrorMessage = $"セッション確認中にエラーが発生しました: {ex.Message}";
                result.CanConnect = true; // エラー時は接続を許可（警告なし）
                return result;
            }
        }

        /// <summary>
        /// ポート生存確認
        /// </summary>
        private async Task<bool> CheckPortAsync(string host, int port, int timeoutMs)
        {
            try
            {
                using (var client = new System.Net.Sockets.TcpClient())
                {
                    var connectTask = client.ConnectAsync(host, port);
                    if (await Task.WhenAny(connectTask, Task.Delay(timeoutMs)) == connectTask)
                    {
                        return client.Connected;
                    }
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// デフォルトOS情報を取得
        /// </summary>
        private OsInfo GetDefaultOsInfo()
        {
            return new OsInfo
            {
                Type = OsType.Unknown,  // Unknownにすることで、WTSを試行する
                MaxSessions = 999,       // 制限なしと仮定
                WarningLevel = WarningLevel.Info,
                IsRdsInstalled = true,   // trueにすることで早期リターンを回避
                OsName = "Unknown",
                OsVersion = "Unknown"
            };
        }
        
        /// <summary>
        /// OS情報を取得（キャッシュ付き）- 旧メソッド、使用しない
        /// </summary>
        [Obsolete("WMIは遅いので使用しない")]
        private async Task<OsInfo> GetOsInfoAsync(string hostName)
        {
            return await Task.Run(() =>
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                
                // キャッシュチェック
                if (_osInfoCache.TryGetValue(hostName, out var cachedInfo) && 
                    _osInfoCacheTime.TryGetValue(hostName, out var cacheTime))
                {
                    if (DateTime.Now - cacheTime < _osInfoCacheTTL)
                    {
                        _debugLog.WriteLine($"[SessionMonitor] Using cached OS info for {hostName} (cached {(DateTime.Now - cacheTime).TotalSeconds:F1}s ago, elapsed: {sw.ElapsedMilliseconds}ms)");
                        return cachedInfo;
                    }
                }

                _debugLog.WriteLine($"[SessionMonitor] OS info not cached or expired, fetching... (elapsed: {sw.ElapsedMilliseconds}ms)");
                var osInfo = new OsInfo();

                try
                {
                    // ローカルマシンの場合
                    if (IsLocalMachine(hostName))
                    {
                        _debugLog.WriteLine($"[SessionMonitor] Getting local OS info... (elapsed: {sw.ElapsedMilliseconds}ms)");
                        osInfo = GetLocalOsInfo();
                        _debugLog.WriteLine($"[SessionMonitor] Local OS info retrieved (elapsed: {sw.ElapsedMilliseconds}ms)");
                    }
                    else
                    {
                        // リモートマシンの場合はWMI経由で取得（タイムアウト付き）
                        _debugLog.WriteLine($"[SessionMonitor] Getting remote OS info via WMI... (elapsed: {sw.ElapsedMilliseconds}ms)");
                        
                        // WMIタイムアウトを短くする（1秒）
                        var wmiTask = Task.Run(() => GetRemoteOsInfo(hostName));
                        if (wmiTask.Wait(TimeSpan.FromSeconds(1)))
                        {
                            osInfo = wmiTask.Result;
                            _debugLog.WriteLine($"[SessionMonitor] Remote OS info retrieved (elapsed: {sw.ElapsedMilliseconds}ms)");
                        }
                        else
                        {
                            _debugLog.WriteLine($"[SessionMonitor] WMI timeout, assuming Workstation (elapsed: {sw.ElapsedMilliseconds}ms)");
                            // タイムアウト時はWorkstationとして扱う
                            osInfo.Type = OsType.Workstation;
                            osInfo.MaxSessions = 1;
                            osInfo.WarningLevel = WarningLevel.Warning;
                        }
                    }
                    
                    // キャッシュに保存
                    _osInfoCache.AddOrUpdate(hostName, osInfo, (k, v) => osInfo);
                    _osInfoCacheTime.AddOrUpdate(hostName, DateTime.Now, (k, v) => DateTime.Now);
                    _debugLog.WriteLine($"[SessionMonitor] OS info cached (total elapsed: {sw.ElapsedMilliseconds}ms)");
                }
                catch (Exception ex)
                {
                    _debugLog.WriteLine($"[SessionMonitor] Failed to get OS info: {ex.Message} (elapsed: {sw.ElapsedMilliseconds}ms)");
                    // エラー時はWorkstationとして扱う（安全側）
                    osInfo.Type = OsType.Workstation;
                    osInfo.MaxSessions = 1;
                    osInfo.WarningLevel = WarningLevel.Warning;
                    osInfo.IsRdsInstalled = false;  // RDS無効として扱うことで早期リターンが効くようにする
                    
                    // エラー時でもキャッシュに保存（無駄な再試行を防ぐ）
                    _osInfoCache.AddOrUpdate(hostName, osInfo, (k, v) => osInfo);
                    _osInfoCacheTime.AddOrUpdate(hostName, DateTime.Now, (k, v) => DateTime.Now);
                }

                return osInfo;
            });
        }

        /// <summary>
        /// ローカルOS情報を取得 - 使用しない
        /// </summary>
        [Obsolete("WMIは遅いので使用しない")]
        private OsInfo GetLocalOsInfo()
        {
            var osInfo = new OsInfo();

            try
            {
                // WMIでOS情報を取得
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        osInfo.OsName = obj["Caption"]?.ToString() ?? "";
                        osInfo.OsVersion = obj["Version"]?.ToString() ?? "";
                        
                        var productType = Convert.ToUInt32(obj["ProductType"]);
                        
                        if (productType == 1)
                        {
                            // ワークステーション
                            osInfo.Type = OsType.Workstation;
                            osInfo.MaxSessions = 1;
                            osInfo.WarningLevel = WarningLevel.Warning;
                        }
                        else
                        {
                            // サーバー
                            osInfo.IsRdsInstalled = CheckRdsInstalled();
                            
                            if (osInfo.IsRdsInstalled)
                            {
                                osInfo.Type = OsType.ServerWithRds;
                                osInfo.MaxSessions = GetRdsLicenseCount();
                                osInfo.WarningLevel = WarningLevel.Info;
                            }
                            else
                            {
                                osInfo.Type = OsType.ServerWithoutRds;
                                osInfo.MaxSessions = 2; // 管理用2セッション
                                osInfo.WarningLevel = WarningLevel.Warning;
                            }
                        }
                        
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                _debugLog.WriteLine($"[SessionMonitor] Error getting local OS info: {ex.Message}");
            }

            return osInfo;
        }

        /// <summary>
        /// リモートOS情報を取得 - 使用しない
        /// </summary>
        [Obsolete("WMIは遅いので使用しない")]
        private OsInfo GetRemoteOsInfo(string hostName)
        {
            var osInfo = new OsInfo();

            try
            {
                var options = new ConnectionOptions
                {
                    Impersonation = ImpersonationLevel.Impersonate,
                    Authentication = AuthenticationLevel.Default
                };

                var scope = new ManagementScope($"\\\\{hostName}\\root\\cimv2", options);
                scope.Connect();

                var query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        osInfo.OsName = obj["Caption"]?.ToString() ?? "";
                        osInfo.OsVersion = obj["Version"]?.ToString() ?? "";
                        
                        var productType = Convert.ToUInt32(obj["ProductType"]);
                        
                        if (productType == 1)
                        {
                            osInfo.Type = OsType.Workstation;
                            osInfo.MaxSessions = 1;
                            osInfo.WarningLevel = WarningLevel.Warning;
                        }
                        else
                        {
                            // リモートのRDS状態確認は別途実装が必要
                            osInfo.IsRdsInstalled = CheckRemoteRdsInstalled(hostName);
                            
                            if (osInfo.IsRdsInstalled)
                            {
                                osInfo.Type = OsType.ServerWithRds;
                                osInfo.MaxSessions = 999; // リモートの場合は正確な数が取れないので大きな値
                                osInfo.WarningLevel = WarningLevel.Info;
                            }
                            else
                            {
                                osInfo.Type = OsType.ServerWithoutRds;
                                osInfo.MaxSessions = 2;
                                osInfo.WarningLevel = WarningLevel.Warning;
                            }
                        }
                        
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                _debugLog.WriteLine($"[SessionMonitor] Error getting remote OS info: {ex.Message}");
            }

            return osInfo;
        }

        /// <summary>
        /// RDSがインストールされているか確認 - 使用しない
        /// </summary>
        [Obsolete("WMIは遅いので使用しない")]
        private bool CheckRdsInstalled()
        {
            try
            {
                // 方法1: レジストリで確認
                using (var key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Terminal Server"))
                {
                    if (key != null)
                    {
                        var tsEnabled = key.GetValue("TSEnabled");
                        if (tsEnabled != null && Convert.ToInt32(tsEnabled) == 1)
                        {
                            // Terminal Server Mode を確認
                            var tsAppCompat = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Terminal Server\TSAppCompat");
                            if (tsAppCompat != null)
                            {
                                // RDSがインストールされている可能性が高い
                                return true;
                            }
                        }
                    }
                }

                // 方法2: WMIで確認
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_TerminalServiceSetting"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var mode = obj["TerminalServerMode"];
                        if (mode != null && Convert.ToInt32(mode) == 1)
                        {
                            return true; // Application Server mode (RDS有効)
                        }
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                _debugLog.WriteLine($"[SessionMonitor] Error checking RDS: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// リモートのRDS状態を確認 - 使用しない
        /// </summary>
        [Obsolete("WMIは遅いので使用しない")]
        private bool CheckRemoteRdsInstalled(string hostName)
        {
            try
            {
                var options = new ConnectionOptions
                {
                    Impersonation = ImpersonationLevel.Impersonate,
                    Authentication = AuthenticationLevel.Default
                };

                var scope = new ManagementScope($"\\\\{hostName}\\root\\cimv2\\TerminalServices", options);
                scope.Connect();

                var query = new ObjectQuery("SELECT * FROM Win32_TerminalServiceSetting");
                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var mode = obj["TerminalServerMode"];
                        if (mode != null && Convert.ToInt32(mode) == 1)
                        {
                            return true;
                        }
                        break;
                    }
                }
            }
            catch
            {
                // エラー時はfalseを返す
            }

            return false;
        }

        /// <summary>
        /// RDSライセンス数を取得 - 使用しない
        /// </summary>
        [Obsolete("WMIは遅いので使用しない")]
        private int GetRdsLicenseCount()
        {
            // 実際のライセンス数取得は複雑なので、ここでは簡易実装
            // 通常は999（実質無制限）として扱う
            return 999;
        }

        /// <summary>
        /// セッション一覧を取得（高速版）
        /// </summary>
        public async Task<SessionInfo[]> GetSessionsAsync(string hostName)
        {
            return await Task.Run(() =>
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                _debugLog.WriteLine($"[SessionMonitor] GetSessionsAsync START for: {hostName}");
                
                // 早期リターン： 明確にWorkstationと判明していて、RDS無効の場合のみスキップ
                // Unknownの場合は必ずWTSを試行する
                if (_osInfoCache.TryGetValue(hostName, out var osInfo))
                {
                    if (osInfo.Type == OsType.Workstation && !osInfo.IsRdsInstalled)
                    {
                        _debugLog.WriteLine($"[SessionMonitor] {hostName}: Known Workstation & RDS disabled -> skip WTS (elapsed: {sw.ElapsedMilliseconds}ms)");
                        return Array.Empty<SessionInfo>();
                    }
                    // ServerWithoutRdsも本当に確実な場合のみスキップ
                    if (osInfo.Type == OsType.ServerWithoutRds && osInfo.OsName != "Unknown")
                    {
                        _debugLog.WriteLine($"[SessionMonitor] {hostName}: Known Server without RDS -> skip WTS (elapsed: {sw.ElapsedMilliseconds}ms)");
                        return Array.Empty<SessionInfo>();
                    }
                }
                
                // ホストフラグを取得
                var flags = _hostFlags.GetOrAdd(hostName, _ => new HostFlags());
                _debugLog.WriteLine($"[SessionMonitor] Host flags retrieved (elapsed: {sw.ElapsedMilliseconds}ms)");
                
                // 負の結果キャッシュ（一定時間は問い合わせ自体をスキップ）
                if (flags.NegativeUntil > DateTime.Now)
                {
                    _debugLog.WriteLine($"[SessionMonitor] {hostName}: negative-cached -> skip until {flags.NegativeUntil:O} (elapsed: {sw.ElapsedMilliseconds}ms)");
                    return Array.Empty<SessionInfo>();
                }
                
                var sessions = new List<SessionInfo>();
                IntPtr serverHandle = IntPtr.Zero;
                bool needClose = false;

                try
                {
                    // ハンドルクリーンアップチェック
                    CleanupOldHandles();
                    
                    // サーバーハンドルを取得（再利用）
                    if (IsLocalMachine(hostName))
                    {
                        serverHandle = (IntPtr)WtsApi32.WTS_CURRENT_SERVER_HANDLE;
                        _debugLog.WriteLine($"[SessionMonitor] Using local server handle (elapsed: {sw.ElapsedMilliseconds}ms)");
                    }
                    else
                    {
                        // ハンドルの再利用
                        if (_serverHandles.TryGetValue(hostName, out serverHandle))
                        {
                            _debugLog.WriteLine($"[SessionMonitor] Found cached handle, validating... (elapsed: {sw.ElapsedMilliseconds}ms)");
                            // ハンドルが有効か簡易チェック
                            uint testCount = 0;
                            IntPtr testPtr = IntPtr.Zero;
                            if (!WtsApi32.WTSEnumerateSessions(serverHandle, 0, 1, out testPtr, out testCount))
                            {
                                // ハンドルが無効になっている
                                _debugLog.WriteLine($"[SessionMonitor] Cached handle invalid, removing (elapsed: {sw.ElapsedMilliseconds}ms)");
                                WtsApi32.WTSCloseServer(serverHandle);
                                _serverHandles.TryRemove(hostName, out _);
                                serverHandle = IntPtr.Zero;
                            }
                            else
                            {
                                _debugLog.WriteLine($"[SessionMonitor] Cached handle valid (elapsed: {sw.ElapsedMilliseconds}ms)");
                                WtsApi32.WTSFreeMemory(testPtr);
                                // ハンドルの最終使用時刻を更新
                                _serverHandleLastUsed.AddOrUpdate(hostName, DateTime.Now, (k, v) => DateTime.Now);
                            }
                        }

                        // 新規接続が必要な場合
                        if (serverHandle == IntPtr.Zero)
                        {
                            _debugLog.WriteLine($"[SessionMonitor] Opening new server connection... (elapsed: {sw.ElapsedMilliseconds}ms)");
                            
                            // FQDN解決を試みる
                            string connectionName = ResolveFqdn(hostName);
                            _debugLog.WriteLine($"[SessionMonitor] Using connection name: {connectionName} (elapsed: {sw.ElapsedMilliseconds}ms)");
                            
                            // WTSOpenServerExをタイムアウト付きで実行（1秒）
                            var openTask = Task.Run(() => WtsApi32.WTSOpenServerEx(connectionName));
                            if (openTask.Wait(TimeSpan.FromSeconds(1)))
                            {
                                serverHandle = openTask.Result;
                                if (serverHandle == IntPtr.Zero)
                                {
                                    _debugLog.WriteLine($"[SessionMonitor] WTSOpenServerEx failed (elapsed: {sw.ElapsedMilliseconds}ms)");
                                    throw new Exception($"Failed to connect to {hostName}");
                                }
                                _debugLog.WriteLine($"[SessionMonitor] Server connection opened (elapsed: {sw.ElapsedMilliseconds}ms)");
                                _serverHandles.AddOrUpdate(hostName, serverHandle, (k, v) => serverHandle);
                                _serverHandleLastUsed.AddOrUpdate(hostName, DateTime.Now, (k, v) => DateTime.Now);
                            }
                            else
                            {
                                _debugLog.WriteLine($"[SessionMonitor] WTSOpenServerEx timeout (elapsed: {sw.ElapsedMilliseconds}ms)");
                                // タイムアウト時は負キャッシュを設定
                                flags.ConsecutiveFailures++;
                                TimeSpan retryInterval = flags.ConsecutiveFailures == 1 ? TimeSpan.FromSeconds(5) :
                                                        flags.ConsecutiveFailures == 2 ? TimeSpan.FromSeconds(10) :
                                                        TimeSpan.FromSeconds(20);
                                flags.NegativeUntil = DateTime.Now.Add(retryInterval);
                                _debugLog.WriteLine($"[SessionMonitor] WTSOpenServerEx timeout (failure #{flags.ConsecutiveFailures}), retry after {retryInterval.TotalSeconds}s");
                                return Array.Empty<SessionInfo>();
                            }
                        }
                    }

                    // WTSEnumerateSessionsExを使用して一括取得（Exを使うかはホストごとに1回だけ判定）
                    IntPtr sessionInfoPtr = IntPtr.Zero;
                    uint sessionCount = 0;
                    uint level = 1; // WTS_SESSION_INFO_1 を使用

                    // Exが使えるかのフラグを確認
                    _debugLog.WriteLine($"[SessionMonitor] Attempting WTSEnumerateSessionsEx (SupportsWtsEx={flags.SupportsWtsEx}) (elapsed: {sw.ElapsedMilliseconds}ms)");
                    if (flags.SupportsWtsEx != false && 
                        WtsApi32.WTSEnumerateSessionsEx(serverHandle, ref level, 0, out sessionInfoPtr, out sessionCount))
                    {
                        try
                        {
                            int sessionInfoSize = Marshal.SizeOf<WtsApi32.WTS_SESSION_INFO_1>();
                            IntPtr currentSession = sessionInfoPtr;

                            for (int i = 0; i < sessionCount; i++)
                            {
                                var wtsSessionInfo = Marshal.PtrToStructure<WtsApi32.WTS_SESSION_INFO_1>(currentSession);
                                
                                var sessionInfo = new SessionInfo
                                {
                                    SessionId = wtsSessionInfo.SessionId,
                                    SessionName = wtsSessionInfo.pSessionName ?? "",
                                    UserName = wtsSessionInfo.pUserName ?? "",
                                    DomainName = wtsSessionInfo.pDomainName ?? "",
                                    State = wtsSessionInfo.State
                                };

                                // 空のセッションは除外
                                if (!string.IsNullOrEmpty(sessionInfo.UserName))
                                {
                                    sessions.Add(sessionInfo);
                                }

                                currentSession = IntPtr.Add(currentSession, sessionInfoSize);
                            }
                        }
                        finally
                        {
                            WtsApi32.WTSFreeMemoryEx(WtsApi32.WTS_TYPE_CLASS.WTSTypeSessionInfoLevel1, sessionInfoPtr, sessionCount);
                        }
                        
                        // Exが成功したら今後も使う
                        _debugLog.WriteLine($"[SessionMonitor] WTSEnumerateSessionsEx succeeded, found {sessionCount} sessions (elapsed: {sw.ElapsedMilliseconds}ms)");
                        flags.SupportsWtsEx = true;
                    }
                    else
                    {
                        // Ex版が失敗した場合は今後使わない
                        if (flags.SupportsWtsEx == null)
                        {
                            _debugLog.WriteLine($"[SessionMonitor] WTSEnumerateSessionsEx failed, marking as unsupported (elapsed: {sw.ElapsedMilliseconds}ms)");
                            flags.SupportsWtsEx = false;
                        }
                        
                        // 従来の方法にフォールバック
                        _debugLog.WriteLine($"[SessionMonitor] Using legacy WTSEnumerateSessions... (elapsed: {sw.ElapsedMilliseconds}ms)");
                        sessions = GetSessionsLegacy(serverHandle, hostName);
                        _debugLog.WriteLine($"[SessionMonitor] Legacy enumeration found {sessions.Count} sessions (elapsed: {sw.ElapsedMilliseconds}ms)");
                    }
                    
                    // 0件の場合も正常な結果として扱う（負キャッシュを設定しない）
                    if (sessions.Count == 0)
                    {
                        _debugLog.WriteLine($"[SessionMonitor] {hostName}: 0 sessions found (elapsed: {sw.ElapsedMilliseconds}ms)");
                        // 成功したので失敗回数をリセット
                        flags.ConsecutiveFailures = 0;
                    }
                    
                    _debugLog.WriteLine($"[SessionMonitor] GetSessionsAsync COMPLETED: returned {sessions.Count} sessions (total elapsed: {sw.ElapsedMilliseconds}ms)");
                }
                catch (Exception ex)
                {
                    _debugLog.WriteLine($"[SessionMonitor] Error enumerating sessions: {ex.Message} (elapsed: {sw.ElapsedMilliseconds}ms)");
                    
                    // RPC到達不能や接続エラーの場合は負キャッシュを設定
                    if (ex.Message.Contains("RPC") || ex.Message.Contains("connect") || 
                        ex.Message.Contains("1722") || ex.Message.Contains("1726"))
                    {
                        flags.ConsecutiveFailures++;
                        TimeSpan retryInterval = flags.ConsecutiveFailures == 1 ? TimeSpan.FromSeconds(5) :
                                                flags.ConsecutiveFailures == 2 ? TimeSpan.FromSeconds(10) :
                                                TimeSpan.FromSeconds(20);
                        flags.NegativeUntil = DateTime.Now.Add(retryInterval);
                        _debugLog.WriteLine($"[SessionMonitor] {hostName}: RPC unreachable (failure #{flags.ConsecutiveFailures}), retry after {retryInterval.TotalSeconds}s");
                    }
                    
                    // エラー時はハンドルを削除（次回再接続）
                    if (!IsLocalMachine(hostName))
                    {
                        _serverHandles.TryRemove(hostName, out _);
                        needClose = true;
                    }
                }
                finally
                {
                    // エラー時のみクローズ（成功時は再利用のため維持）
                    if (needClose && serverHandle != IntPtr.Zero && serverHandle != (IntPtr)WtsApi32.WTS_CURRENT_SERVER_HANDLE)
                    {
                        WtsApi32.WTSCloseServer(serverHandle);
                    }
                }

                return sessions.ToArray();
            });
        }

        /// <summary>
        /// レガシー版のセッション取得（フォールバック用）
        /// </summary>
        private List<SessionInfo> GetSessionsLegacy(IntPtr serverHandle, string hostName)
        {
            var sessions = new List<SessionInfo>();
            IntPtr sessionInfoPtr = IntPtr.Zero;
            uint sessionCount = 0;

            if (WtsApi32.WTSEnumerateSessions(serverHandle, 0, 1, out sessionInfoPtr, out sessionCount))
            {
                try
                {
                    int sessionInfoSize = Marshal.SizeOf<WtsApi32.WTS_SESSION_INFO>();
                    IntPtr currentSession = sessionInfoPtr;

                    for (int i = 0; i < sessionCount; i++)
                    {
                        var wtsSessionInfo = Marshal.PtrToStructure<WtsApi32.WTS_SESSION_INFO>(currentSession);
                        
                        var sessionInfo = new SessionInfo
                        {
                            SessionId = wtsSessionInfo.SessionId,
                            SessionName = wtsSessionInfo.pWinStationName ?? "",
                            State = wtsSessionInfo.State
                        };

                        // ユーザー名を取得
                        if (QuerySessionString(serverHandle, wtsSessionInfo.SessionId, 
                            WtsApi32.WTS_INFO_CLASS.WTSUserName, out string userName))
                        {
                            sessionInfo.UserName = userName;
                        }

                        // ドメイン名を取得
                        if (QuerySessionString(serverHandle, wtsSessionInfo.SessionId,
                            WtsApi32.WTS_INFO_CLASS.WTSDomainName, out string domainName))
                        {
                            sessionInfo.DomainName = domainName;
                        }

                        // 空のセッションは除外
                        if (!string.IsNullOrEmpty(sessionInfo.UserName))
                        {
                            sessions.Add(sessionInfo);
                        }

                        currentSession = IntPtr.Add(currentSession, sessionInfoSize);
                    }
                }
                finally
                {
                    WtsApi32.WTSFreeMemory(sessionInfoPtr);
                }
            }
            return sessions;
        }

        /// <summary>
        /// セッション情報文字列を取得
        /// </summary>
        private bool QuerySessionString(IntPtr serverHandle, uint sessionId, 
            WtsApi32.WTS_INFO_CLASS infoClass, out string result)
        {
            result = string.Empty;
            IntPtr buffer = IntPtr.Zero;
            uint bytesReturned = 0;

            try
            {
                if (WtsApi32.WTSQuerySessionInformation(serverHandle, sessionId, infoClass, 
                    out buffer, out bytesReturned))
                {
                    result = Marshal.PtrToStringAuto(buffer) ?? string.Empty;
                    return true;
                }
            }
            finally
            {
                if (buffer != IntPtr.Zero)
                {
                    WtsApi32.WTSFreeMemory(buffer);
                }
            }

            return false;
        }

        /// <summary>
        /// ローカルマシンかどうか判定
        /// </summary>
        private bool IsLocalMachine(string hostName)
        {
            if (string.IsNullOrEmpty(hostName))
                return true;

            var localNames = new[] { "localhost", "127.0.0.1", ".", 
                Environment.MachineName, "::1" };

            return localNames.Any(name => 
                hostName.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 他のユーザーが使用中か判定
        /// </summary>
        private bool IsInUseByOthers(SessionInfo[] sessions, string currentUser)
        {
            if (sessions == null || sessions.Length == 0)
                return false;

            // アクティブなセッションで、現在のユーザー以外のものがあるか
            return sessions.Any(s => 
                s.IsActive && 
                !s.UserName.Equals(currentUser, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 警告メッセージを生成
        /// </summary>
        private string GenerateWarningMessage(SessionCheckResult result)
        {
            if (!result.IsInUseByOthers)
                return string.Empty;

            var otherUsers = result.Sessions
                .Where(s => s.IsActive && 
                    !s.UserName.Equals(result.CurrentUser, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (otherUsers.Length == 0)
                return string.Empty;

            var sb = new StringBuilder();

            switch (result.OsInfo.Type)
            {
                case OsType.Workstation:
                    sb.AppendLine("⚠️ 警告: 他のユーザーが使用中です");
                    sb.AppendLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                    foreach (var user in otherUsers)
                    {
                        sb.AppendLine($"{user.FullUserName} が現在使用中です。");
                    }
                    sb.AppendLine();
                    sb.AppendLine("接続すると、既存のセッションが強制的に");
                    sb.AppendLine("切断されます。");
                    break;

                case OsType.ServerWithoutRds:
                    sb.AppendLine("⚠️ 警告: 管理用セッション制限");
                    sb.AppendLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                    sb.AppendLine("現在のアクティブセッション:");
                    foreach (var user in otherUsers)
                    {
                        sb.AppendLine($"• {user.FullUserName} ({user.SessionName}) - {user.State}");
                    }
                    sb.AppendLine();
                    sb.AppendLine("このサーバーはRDSがインストールされていないため、");
                    sb.AppendLine("管理用接続（2セッション）のみ利用可能です。");
                    sb.AppendLine();
                    sb.AppendLine("接続すると、既存のセッションが切断される");
                    sb.AppendLine("可能性があります。");
                    break;

                case OsType.ServerWithRds:
                    sb.AppendLine("ℹ️ 情報: 複数のユーザーが接続中です");
                    sb.AppendLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                    sb.AppendLine("現在のアクティブセッション:");
                    foreach (var session in result.Sessions.Where(s => !string.IsNullOrEmpty(s.UserName)))
                    {
                        sb.AppendLine($"• {session.FullUserName} ({session.SessionName}) - {session.State}");
                    }
                    sb.AppendLine();
                    sb.AppendLine("このサーバーは複数同時接続をサポートしています。");
                    sb.AppendLine("接続しても既存セッションは維持されます。");
                    break;
            }

            return sb.ToString();
        }

        /// <summary>
        /// FQDNを解決（キャッシュ付き）
        /// </summary>
        private string ResolveFqdn(string hostName)
        {
            // キャッシュチェック
            if (_fqdnCache.TryGetValue(hostName, out var cachedFqdn) &&
                _fqdnCacheTime.TryGetValue(hostName, out var cacheTime))
            {
                if (DateTime.Now - cacheTime < _fqdnCacheTTL)
                {
                    _debugLog.WriteLine($"[SessionMonitor] Using cached FQDN: {cachedFqdn}");
                    return cachedFqdn;
                }
            }
            
            string fqdn = hostName;
            
            try
            {
                // IPアドレスかどうかチェック
                if (System.Net.IPAddress.TryParse(hostName, out _))
                {
                    // IPアドレスの場合はそのまま使用
                    fqdn = hostName;
                }
                else
                {
                    // ホスト名の場合はFQDNを取得
                    var hostEntry = System.Net.Dns.GetHostEntry(hostName);
                    fqdn = hostEntry.HostName;
                    _debugLog.WriteLine($"[SessionMonitor] Resolved FQDN: {hostName} -> {fqdn}");
                }
            }
            catch (Exception ex)
            {
                _debugLog.WriteLine($"[SessionMonitor] Failed to resolve FQDN for {hostName}: {ex.Message}");
                // 解決できない場合は元のホスト名を使用
                fqdn = hostName;
            }
            
            // キャッシュに保存
            _fqdnCache.AddOrUpdate(hostName, fqdn, (k, v) => fqdn);
            _fqdnCacheTime.AddOrUpdate(hostName, DateTime.Now, (k, v) => DateTime.Now);
            
            return fqdn;
        }
        
        /// <summary>
        /// 古いハンドルをクリーンアップ
        /// </summary>
        private void CleanupOldHandles()
        {
            // クリーンアップインターバルチェック
            if (DateTime.Now - _lastHandleCleanup < _handleCleanupInterval)
                return;
                
            _lastHandleCleanup = DateTime.Now;
            _debugLog.WriteLine("[SessionMonitor] Starting handle cleanup...");
            
            var handlersToRemove = new List<string>();
            var now = DateTime.Now;
            
            foreach (var kvp in _serverHandleLastUsed)
            {
                if (now - kvp.Value > _handleKeepAlive)
                {
                    handlersToRemove.Add(kvp.Key);
                }
            }
            
            foreach (var hostName in handlersToRemove)
            {
                if (_serverHandles.TryRemove(hostName, out var handle))
                {
                    if (handle != IntPtr.Zero && handle != (IntPtr)WtsApi32.WTS_CURRENT_SERVER_HANDLE)
                    {
                        WtsApi32.WTSCloseServer(handle);
                        _debugLog.WriteLine($"[SessionMonitor] Closed old handle for {hostName} (idle for {(now - _serverHandleLastUsed[hostName]).TotalMinutes:F1} minutes)");
                    }
                    _serverHandleLastUsed.TryRemove(hostName, out _);
                }
            }
            
            if (handlersToRemove.Count > 0)
            {
                _debugLog.WriteLine($"[SessionMonitor] Cleaned up {handlersToRemove.Count} old handles");
            }
        }
        
        /// <summary>
        /// 接続可能か判定
        /// </summary>
        private bool DetermineConnectability(SessionCheckResult result)
        {
            // エラーがある場合は接続可能（警告なしで進む）
            if (!result.IsSuccess)
                return true;

            // 他のユーザーが使用していない場合は接続可能
            if (!result.IsInUseByOthers)
                return true;

            // RDS有りのサーバーの場合は常に接続可能
            if (result.OsInfo.Type == OsType.ServerWithRds)
                return true;

            // それ以外（ワークステーションやRDS無しサーバー）は警告後に判断
            return false;
        }
    }
}