using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using RemoteWakeConnect.Models;

namespace RemoteWakeConnect.Services
{
    public class RemoteDesktopService
    {
        private readonly RdpFileService _rdpFileService;
        private readonly NetworkService _networkService;
        private readonly ConnectionHistoryService _connectionHistoryService;
        
        public RemoteDesktopService()
        {
            _rdpFileService = new RdpFileService();
            _networkService = new NetworkService();
            _connectionHistoryService = new ConnectionHistoryService();
        }

        public async Task<bool> ConnectAsync(RdpConnection connection)
        {
            try
            {
                string rdpFilePath;
                bool isTemporaryFile = false;
                
                // 既存のRDPファイルパスがある場合はそれを使用
                System.Diagnostics.Debug.WriteLine($"[RemoteDesktopService] RdpFilePath check: {connection.RdpFilePath}");
                System.Diagnostics.Debug.WriteLine($"[RemoteDesktopService] File exists: {(!string.IsNullOrEmpty(connection.RdpFilePath) && File.Exists(connection.RdpFilePath))}");
                
                if (!string.IsNullOrEmpty(connection.RdpFilePath) && File.Exists(connection.RdpFilePath))
                {
                    rdpFilePath = connection.RdpFilePath;
                    System.Diagnostics.Debug.WriteLine($"[RemoteDesktopService] Using existing file: {rdpFilePath}");
                    // 既存ファイルを最新の設定で更新
                    _rdpFileService.SaveRdpFile(rdpFilePath, connection);
                }
                else
                {
                    // RDPファイル名を決定（元のファイル名があればそれを使用、なければ生成）
                    string rdpFileName;
                    if (!string.IsNullOrEmpty(connection.Name))
                    {
                        // 元のRDPファイル名をそのまま使用
                        rdpFileName = Path.GetFileName(connection.Name);
                    }
                    else if (!string.IsNullOrEmpty(connection.ComputerName))
                    {
                        // コンピュータ名を使用
                        rdpFileName = $"{connection.ComputerName}.rdp";
                    }
                    else if (!string.IsNullOrEmpty(connection.FullAddress))
                    {
                        // フルアドレスから安全なファイル名を生成
                        var safeName = connection.FullAddress.Replace(":", "_").Replace(".", "_");
                        rdpFileName = $"{safeName}.rdp";
                    }
                    else
                    {
                        // デフォルト名
                        rdpFileName = "RemoteWakeConnect.rdp";
                    }
                    
                    // 一時的なRDPファイルを作成（固定の場所に固定の名前で）
                    rdpFilePath = Path.Combine(Path.GetTempPath(), rdpFileName);
                    isTemporaryFile = true;
                    
                    // RDPファイルを保存
                    _rdpFileService.SaveRdpFile(rdpFilePath, connection);
                }

                try
                {
                    // mstscコマンドを実行
                    var processInfo = new ProcessStartInfo
                    {
                        FileName = "mstsc.exe",
                        Arguments = $"\"{rdpFilePath}\"",
                        UseShellExecute = false,
                        CreateNoWindow = false
                    };

                    var process = Process.Start(processInfo);
                    
                    // プロセスが開始されるまで待機
                    await Task.Delay(1000);
                    
                    // 接続履歴を更新
                    connection.LastConnection = DateTime.Now;
                    
                    // 非同期でMACアドレスを取得して保存
                    _ = Task.Run(async () =>
                    {
                        // 接続が確立されるまで待つ（より長く待機）
                        await Task.Delay(3000);
                        await UpdateConnectionWithMacAddressAsync(connection);
                        
                        // 失敗した場合はもう一度試す
                        await Task.Delay(2000);
                        await UpdateConnectionWithMacAddressAsync(connection);
                    });
                    
                    return process != null && !process.HasExited;
                }
                finally
                {
                    // 一時ファイルの場合のみ削除を遅延実行
                    if (isTemporaryFile)
                    {
                        _ = Task.Run(async () =>
                        {
                            await Task.Delay(5000);
                            try
                            {
                                if (File.Exists(rdpFilePath))
                                {
                                    File.Delete(rdpFilePath);
                                }
                            }
                            catch { }
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"リモートデスクトップ接続に失敗しました: {ex.Message}", ex);
            }
        }

        public bool IsRemoteDesktopAvailable()
        {
            try
            {
                var mstscPath = Path.Combine(Environment.SystemDirectory, "mstsc.exe");
                return File.Exists(mstscPath);
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> TestConnectionAsync(string hostNameOrAddress, int port = 3389)
        {
            try
            {
                using (var client = new System.Net.Sockets.TcpClient())
                {
                    var connectTask = client.ConnectAsync(hostNameOrAddress, port);
                    var timeoutTask = Task.Delay(5000);
                    
                    var completedTask = await Task.WhenAny(connectTask, timeoutTask);
                    
                    if (completedTask == timeoutTask)
                    {
                        return false;
                    }
                    
                    await connectTask;
                    return client.Connected;
                }
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// 接続後にMACアドレスを取得して履歴を更新
        /// </summary>
        private async Task UpdateConnectionWithMacAddressAsync(RdpConnection connection)
        {
            try
            {
                // 履歴内の接続情報を検索
                var existingConnection = _connectionHistoryService.FindByAddress(connection.FullAddress);
                if (existingConnection == null)
                {
                    System.Diagnostics.Debug.WriteLine($"履歴に接続情報が見つかりません: {connection.FullAddress}");
                    return;
                }
                
                // MACアドレスが未取得の場合のみ取得を試みる
                if (string.IsNullOrEmpty(existingConnection.MacAddress))
                {
                    string hostNameOrAddress = !string.IsNullOrEmpty(connection.ComputerName) 
                        ? connection.ComputerName 
                        : (!string.IsNullOrEmpty(connection.IpAddressValue) 
                            ? connection.IpAddressValue 
                            : connection.FullAddress.Split(':')[0]);
                    
                    System.Diagnostics.Debug.WriteLine($"MACアドレス取得開始: {hostNameOrAddress}");
                    
                    if (!string.IsNullOrEmpty(hostNameOrAddress))
                    {
                        // MACアドレスを取得（複数の方法を試す）
                        var macAddress = await _networkService.GetMacAddressAsync(hostNameOrAddress);
                        
                        // 最初の方法で失敗した場合、nbtstatも試す
                        if (string.IsNullOrEmpty(macAddress))
                        {
                            System.Diagnostics.Debug.WriteLine("arp/WMI失敗、nbtstatを試します");
                            macAddress = await _networkService.GetMacFromNbtstatAsync(hostNameOrAddress);
                        }
                        
                        if (!string.IsNullOrEmpty(macAddress))
                        {
                            System.Diagnostics.Debug.WriteLine($"MACアドレス取得成功: {macAddress}");
                            existingConnection.MacAddress = macAddress;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("MACアドレス取得失敗");
                        }
                    }
                }
                
                // ユーザー名を更新（空の場合のみ）
                if (string.IsNullOrEmpty(existingConnection.Username) && !string.IsNullOrEmpty(connection.Username))
                {
                    existingConnection.Username = connection.Username;
                }
                
                // IPアドレスも更新（コンピュータ名で接続した場合）
                if (string.IsNullOrEmpty(existingConnection.IpAddressValue) && 
                    !string.IsNullOrEmpty(connection.ComputerName))
                {
                    try
                    {
                        var hostEntry = await System.Net.Dns.GetHostEntryAsync(connection.ComputerName);
                        var ipv4Address = hostEntry.AddressList
                            .FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                        if (ipv4Address != null)
                        {
                            existingConnection.IpAddressValue = ipv4Address.ToString();
                            System.Diagnostics.Debug.WriteLine($"IPアドレス解決成功: {ipv4Address}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"IPアドレス解決失敗: {ex.Message}");
                    }
                }
                
                // 履歴を更新して保存
                _connectionHistoryService.UpdateConnection(existingConnection);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MACアドレス取得エラー: {ex.Message}");
            }
        }
    }
}