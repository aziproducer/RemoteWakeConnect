using System;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RemoteWakeConnect.Services
{
    public class WakeOnLanService
    {
        private const int WOL_PORT = 9;

        public async Task<bool> SendMagicPacketAsync(string macAddress, string? ipAddress = null)
        {
            try
            {
                var cleanMacAddress = CleanMacAddress(macAddress);
                if (!IsValidMacAddress(cleanMacAddress))
                {
                    throw new ArgumentException("無効なMACアドレス形式です。");
                }

                byte[] magicPacket = BuildMagicPacket(cleanMacAddress);
                
                using (var client = new UdpClient())
                {
                    client.EnableBroadcast = true;
                    
                    // ブロードキャストアドレスを決定
                    IPEndPoint endPoint;
                    if (string.IsNullOrEmpty(ipAddress))
                    {
                        endPoint = new IPEndPoint(IPAddress.Broadcast, WOL_PORT);
                    }
                    else
                    {
                        // IPアドレスかコンピュータ名かを判定
                        IPAddress? ip = null;
                        if (IPAddress.TryParse(ipAddress, out var parsedIp))
                        {
                            ip = parsedIp;
                        }
                        else
                        {
                            // コンピュータ名の場合はDNS解決を試みる
                            try
                            {
                                var hostEntry = await Dns.GetHostEntryAsync(ipAddress);
                                ip = hostEntry.AddressList.FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork);
                            }
                            catch
                            {
                                // DNS解決に失敗した場合はブロードキャストを使用
                                endPoint = new IPEndPoint(IPAddress.Broadcast, WOL_PORT);
                                await client.SendAsync(magicPacket, magicPacket.Length, endPoint);
                                return true;
                            }
                        }
                        
                        if (ip != null)
                        {
                            var broadcastAddress = GetBroadcastAddress(ip);
                            endPoint = new IPEndPoint(broadcastAddress, WOL_PORT);
                        }
                        else
                        {
                            endPoint = new IPEndPoint(IPAddress.Broadcast, WOL_PORT);
                        }
                    }

                    // マジックパケットを送信
                    await client.SendAsync(magicPacket, magicPacket.Length, endPoint);
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Wake On LANパケットの送信に失敗しました: {ex.Message}", ex);
            }
        }

        public async Task<bool> PingHostAsync(string hostNameOrAddress, int timeout = 5000)
        {
            try
            {
                using (var ping = new Ping())
                {
                    PingReply reply = await ping.SendPingAsync(hostNameOrAddress);
                    return reply.Status == IPStatus.Success;
                }
            }
            catch
            {
                return false;
            }
        }

        private string CleanMacAddress(string macAddress)
        {
            return macAddress.Replace(":", "").Replace("-", "").Replace(" ", "").ToUpper();
        }

        private bool IsValidMacAddress(string macAddress)
        {
            return Regex.IsMatch(macAddress, "^[0-9A-F]{12}$");
        }

        private byte[] BuildMagicPacket(string macAddress)
        {
            byte[] packet = new byte[102];
            
            // 最初の6バイトは0xFF
            for (int i = 0; i < 6; i++)
            {
                packet[i] = 0xFF;
            }

            // MACアドレスをバイト配列に変換
            byte[] macBytes = new byte[6];
            for (int i = 0; i < 6; i++)
            {
                macBytes[i] = Convert.ToByte(macAddress.Substring(i * 2, 2), 16);
            }

            // MACアドレスを16回繰り返す
            for (int i = 1; i <= 16; i++)
            {
                Array.Copy(macBytes, 0, packet, i * 6, 6);
            }

            return packet;
        }

        private IPAddress GetBroadcastAddress(IPAddress address)
        {
            // 簡易的な実装：最後のオクテットを255にする
            byte[] ipBytes = address.GetAddressBytes();
            ipBytes[3] = 255;
            return new IPAddress(ipBytes);
        }
    }
}