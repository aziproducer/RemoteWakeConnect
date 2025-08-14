using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using RemoteWakeConnect.Models;

namespace RemoteWakeConnect.Services
{
    public class RemoteDesktopService
    {
        private readonly RdpFileService _rdpFileService;
        
        public RemoteDesktopService()
        {
            _rdpFileService = new RdpFileService();
        }

        public async Task<bool> ConnectAsync(RdpConnection connection)
        {
            try
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
                string tempRdpFile = Path.Combine(Path.GetTempPath(), rdpFileName);
                
                try
                {
                    // RDPファイルを保存
                    _rdpFileService.SaveRdpFile(tempRdpFile, connection);

                    // mstscコマンドを実行
                    var processInfo = new ProcessStartInfo
                    {
                        FileName = "mstsc.exe",
                        Arguments = $"\"{tempRdpFile}\"",
                        UseShellExecute = false,
                        CreateNoWindow = false
                    };

                    var process = Process.Start(processInfo);
                    
                    // プロセスが開始されるまで待機
                    await Task.Delay(1000);
                    
                    // 接続履歴を更新
                    connection.LastConnection = DateTime.Now;
                    
                    return process != null && !process.HasExited;
                }
                finally
                {
                    // 一時ファイルの削除を遅延実行
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(5000);
                        try
                        {
                            if (File.Exists(tempRdpFile))
                            {
                                File.Delete(tempRdpFile);
                            }
                        }
                        catch { }
                    });
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
    }
}