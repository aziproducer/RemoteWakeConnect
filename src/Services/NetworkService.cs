using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RemoteWakeConnect.Services
{
    public class NetworkService
    {
        /// <summary>
        /// IPアドレスまたはホスト名からMACアドレスを取得
        /// </summary>
        public async Task<string?> GetMacAddressAsync(string hostNameOrAddress)
        {
            try
            {
                // まずPingを送信してARPテーブルに登録
                await PingHostAsync(hostNameOrAddress);
                
                // IPアドレスを取得
                string ipAddress = await ResolveIpAddressAsync(hostNameOrAddress);
                if (string.IsNullOrEmpty(ipAddress))
                    return null;

                // ARPテーブルからMACアドレスを取得
                string? macAddress = GetMacFromArpTable(ipAddress);
                if (!string.IsNullOrEmpty(macAddress))
                    return macAddress;

                // WMIを使用して取得を試みる（同一ネットワーク内のWindowsマシンの場合）
                macAddress = GetMacFromWmi(hostNameOrAddress);
                if (!string.IsNullOrEmpty(macAddress))
                    return macAddress;

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"MACアドレス取得エラー: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// ホスト名からIPアドレスを解決
        /// </summary>
        private async Task<string> ResolveIpAddressAsync(string hostNameOrAddress)
        {
            try
            {
                // すでにIPアドレスの場合はそのまま返す
                if (IPAddress.TryParse(hostNameOrAddress, out var ip))
                    return ip.ToString();

                // ホスト名からIPアドレスを解決
                var hostEntry = await Dns.GetHostEntryAsync(hostNameOrAddress);
                var ipv4Address = hostEntry.AddressList
                    .FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                
                return ipv4Address?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// ホストにPingを送信
        /// </summary>
        private async Task<bool> PingHostAsync(string hostNameOrAddress)
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = await ping.SendPingAsync(hostNameOrAddress);
                    return reply.Status == IPStatus.Success;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// ARPテーブルからMACアドレスを取得
        /// </summary>
        private string? GetMacFromArpTable(string ipAddress)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "arp",
                        Arguments = "-a",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true,
                        StandardOutputEncoding = System.Text.Encoding.Default
                    }
                };

                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                // ARPテーブルを解析
                // 例: "192.168.1.100     00-11-22-33-44-55     動的"
                var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (line.Contains(ipAddress))
                    {
                        // MACアドレスのパターンをマッチング
                        var macPattern = @"([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})";
                        var match = Regex.Match(line, macPattern);
                        if (match.Success)
                        {
                            // ハイフンをコロンに変換して返す
                            return match.Value.Replace('-', ':').ToUpper();
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ARPテーブル取得エラー: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// WMIを使用してMACアドレスを取得（リモートWindowsマシン用）
        /// </summary>
        private string? GetMacFromWmi(string hostName)
        {
            try
            {
                // ローカルマシンの場合
                if (hostName.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                    hostName.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase) ||
                    hostName.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase))
                {
                    using (var searcher = new ManagementObjectSearcher(
                        "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"))
                    {
                        foreach (ManagementObject obj in searcher.Get())
                        {
                            var macAddress = obj["MACAddress"]?.ToString();
                            if (!string.IsNullOrEmpty(macAddress))
                                return macAddress.ToUpper();
                        }
                    }
                }
                // リモートマシンの場合（認証が必要な場合があるため、エラーになる可能性が高い）
                else
                {
                    var scope = new ManagementScope($@"\\{hostName}\root\cimv2");
                    scope.Connect();
                    
                    using (var searcher = new ManagementObjectSearcher(scope,
                        new ObjectQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")))
                    {
                        foreach (ManagementObject obj in searcher.Get())
                        {
                            var macAddress = obj["MACAddress"]?.ToString();
                            if (!string.IsNullOrEmpty(macAddress))
                                return macAddress.ToUpper();
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WMI取得エラー: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// nbtstatコマンドを使用してMACアドレスを取得（NetBIOS名が有効な場合）
        /// </summary>
        public async Task<string?> GetMacFromNbtstatAsync(string hostNameOrAddress)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "nbtstat",
                        Arguments = $"-a {hostNameOrAddress}",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true,
                        StandardOutputEncoding = System.Text.Encoding.Default
                    }
                };

                process.Start();
                
                var outputTask = Task.Run(() => process.StandardOutput.ReadToEnd());
                if (await Task.WhenAny(outputTask, Task.Delay(5000)) == outputTask)
                {
                    string output = await outputTask;
                    process.WaitForExit();

                    // MACアドレスを探す
                    // 例: "MACアドレス = 00-11-22-33-44-55"
                    var macPattern = @"(?:MAC|Mac).*?([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})";
                    var match = Regex.Match(output, macPattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        var macMatch = Regex.Match(match.Value, @"([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})");
                        if (macMatch.Success)
                        {
                            return macMatch.Value.Replace('-', ':').ToUpper();
                        }
                    }
                }
                else
                {
                    try { process.Kill(); } catch { }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"nbtstat取得エラー: {ex.Message}");
                return null;
            }
        }
    }
}