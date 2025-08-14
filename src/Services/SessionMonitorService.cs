using System;
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
    /// セッション監視サービス
    /// </summary>
    public class SessionMonitorService
    {
        private readonly DebugLogService _debugLog;

        public SessionMonitorService()
        {
            _debugLog = new DebugLogService();
        }

        /// <summary>
        /// セッション情報を確認
        /// </summary>
        /// <param name="hostName">ホスト名またはIPアドレス</param>
        /// <returns>セッション確認結果</returns>
        public async Task<SessionCheckResult> CheckSessionsAsync(string hostName)
        {
            var result = new SessionCheckResult
            {
                CurrentUser = Environment.UserName
            };

            try
            {
                _debugLog.WriteLine($"[SessionMonitor] Checking sessions for: {hostName}");

                // OS情報を取得
                result.OsInfo = await GetOsInfoAsync(hostName);
                _debugLog.WriteLine($"[SessionMonitor] OS Type: {result.OsInfo.Type}, RDS: {result.OsInfo.IsRdsInstalled}");

                // セッション情報を取得
                result.Sessions = await GetSessionsAsync(hostName);
                _debugLog.WriteLine($"[SessionMonitor] Found {result.Sessions.Length} sessions");

                // 他のユーザーが使用中か判定
                result.IsInUseByOthers = IsInUseByOthers(result.Sessions, result.CurrentUser);

                // 警告メッセージを生成
                result.WarningMessage = GenerateWarningMessage(result);

                // 接続可能か判定
                result.CanConnect = DetermineConnectability(result);

                return result;
            }
            catch (Exception ex)
            {
                _debugLog.WriteLine($"[SessionMonitor] Error: {ex.Message}");
                result.ErrorMessage = $"セッション確認中にエラーが発生しました: {ex.Message}";
                result.CanConnect = true; // エラー時は接続を許可（警告なし）
                return result;
            }
        }

        /// <summary>
        /// OS情報を取得
        /// </summary>
        private async Task<OsInfo> GetOsInfoAsync(string hostName)
        {
            return await Task.Run(() =>
            {
                var osInfo = new OsInfo();

                try
                {
                    // ローカルマシンの場合
                    if (IsLocalMachine(hostName))
                    {
                        osInfo = GetLocalOsInfo();
                    }
                    else
                    {
                        // リモートマシンの場合はWMI経由で取得
                        osInfo = GetRemoteOsInfo(hostName);
                    }
                }
                catch (Exception ex)
                {
                    _debugLog.WriteLine($"[SessionMonitor] Failed to get OS info: {ex.Message}");
                    // エラー時はWorkstationとして扱う（安全側）
                    osInfo.Type = OsType.Workstation;
                    osInfo.MaxSessions = 1;
                    osInfo.WarningLevel = WarningLevel.Warning;
                }

                return osInfo;
            });
        }

        /// <summary>
        /// ローカルOS情報を取得
        /// </summary>
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
        /// リモートOS情報を取得
        /// </summary>
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
        /// RDSがインストールされているか確認
        /// </summary>
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
        /// リモートのRDS状態を確認
        /// </summary>
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
        /// RDSライセンス数を取得
        /// </summary>
        private int GetRdsLicenseCount()
        {
            // 実際のライセンス数取得は複雑なので、ここでは簡易実装
            // 通常は999（実質無制限）として扱う
            return 999;
        }

        /// <summary>
        /// セッション一覧を取得
        /// </summary>
        public async Task<SessionInfo[]> GetSessionsAsync(string hostName)
        {
            return await Task.Run(() =>
            {
                var sessions = new List<SessionInfo>();
                IntPtr serverHandle = IntPtr.Zero;

                try
                {
                    // サーバーハンドルを開く
                    if (IsLocalMachine(hostName))
                    {
                        serverHandle = (IntPtr)WtsApi32.WTS_CURRENT_SERVER_HANDLE;
                    }
                    else
                    {
                        serverHandle = WtsApi32.WTSOpenServerEx(hostName);
                        if (serverHandle == IntPtr.Zero)
                        {
                            throw new Exception($"Failed to connect to {hostName}");
                        }
                    }

                    // セッション一覧を取得
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
                }
                catch (Exception ex)
                {
                    _debugLog.WriteLine($"[SessionMonitor] Error enumerating sessions: {ex.Message}");
                }
                finally
                {
                    if (serverHandle != IntPtr.Zero && serverHandle != (IntPtr)WtsApi32.WTS_CURRENT_SERVER_HANDLE)
                    {
                        WtsApi32.WTSCloseServer(serverHandle);
                    }
                }

                return sessions.ToArray();
            });
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