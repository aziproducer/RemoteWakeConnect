using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using RemoteWakeConnect.Models;

namespace RemoteWakeConnect.Services
{
    public class RdpFileService
    {
        public RdpConnection LoadRdpFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"RDPファイルが見つかりません: {filePath}");
            }

            var connection = new RdpConnection();
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            
            foreach (var line in lines)
            {
                ParseRdpLine(line, connection);
            }

            // ファイル名から接続名を設定
            if (string.IsNullOrEmpty(connection.Name))
            {
                connection.Name = Path.GetFileNameWithoutExtension(filePath);
            }

            return connection;
        }

        public void SaveRdpFile(string filePath, RdpConnection connection)
        {
            var lines = new List<string>();

            // 基本設定
            lines.Add($"full address:s:{connection.FullAddress}");
            lines.Add($"username:s:{connection.Username}");
            
            if (!string.IsNullOrEmpty(connection.Domain))
            {
                lines.Add($"domain:s:{connection.Domain}");
            }

            // 画面設定
            lines.Add($"screen mode id:i:{connection.ScreenModeId}");
            lines.Add($"use multimon:i:{(connection.UseMultimon ? 1 : 0)}");
            
            if (connection.UseMultimon && connection.SelectedMonitors > 0)
            {
                lines.Add($"selectedmonitors:s:{BuildSelectedMonitorsString(connection.SelectedMonitors)}");
            }

            lines.Add($"desktopwidth:i:{connection.DesktopWidth}");
            lines.Add($"desktopheight:i:{connection.DesktopHeight}");

            // エクスペリエンス設定
            lines.Add($"session bpp:i:{connection.ColorDepth}");
            lines.Add("compression:i:1");
            lines.Add($"connection type:i:{connection.ConnectionType}");
            lines.Add($"networkautodetect:i:{(connection.ConnectionType == 0 ? 1 : 0)}");
            lines.Add($"bandwidthautodetect:i:{(connection.ConnectionType == 0 ? 1 : 0)}");
            lines.Add("displayconnectionbar:i:1");
            lines.Add("enableworkspacereconnect:i:0");
            
            // 表示設定
            lines.Add($"disable wallpaper:i:{(connection.DesktopBackground ? 0 : 1)}");
            lines.Add($"allow font smoothing:i:{(connection.FontSmoothing ? 1 : 0)}");
            lines.Add($"allow desktop composition:i:{(connection.DesktopComposition ? 1 : 0)}");
            lines.Add($"disable full window drag:i:{(connection.ShowWindowContents ? 0 : 1)}");
            lines.Add($"disable menu anims:i:{(connection.MenuAnimations ? 0 : 1)}");
            lines.Add($"disable themes:i:{(connection.VisualStyles ? 0 : 1)}");
            lines.Add("disable cursor setting:i:0");
            lines.Add($"bitmapcachepersistenable:i:{(connection.BitmapCaching ? 1 : 0)}");
            
            // オーディオ設定
            lines.Add($"audiomode:i:{connection.AudioMode}");
            lines.Add($"audiocapturemode:i:{(connection.AudioRecord ? 1 : 0)}");
            lines.Add("videoplaybackmode:i:1");
            
            // キーボード設定
            lines.Add($"keyboardhook:i:{connection.KeyboardMode}");
            
            // ローカルリソース設定
            lines.Add($"redirectprinters:i:{(connection.RedirectPrinters ? 1 : 0)}");
            lines.Add($"redirectcomports:i:{(connection.RedirectPorts ? 1 : 0)}");
            lines.Add($"redirectsmartcards:i:{(connection.RedirectSmartCards ? 1 : 0)}");
            lines.Add($"redirectclipboard:i:{(connection.RedirectClipboard ? 1 : 0)}");
            lines.Add($"redirectdrives:i:{(connection.RedirectDrives ? 1 : 0)}");
            lines.Add($"redirectposdevices:i:{(connection.RedirectPnpDevices ? 1 : 0)}");
            
            // 再接続設定
            lines.Add($"autoreconnection enabled:i:{(connection.AutoReconnect ? 1 : 0)}");
            lines.Add("authentication level:i:2");
            lines.Add("prompt for credentials:i:0");
            lines.Add("negotiate security layer:i:1");
            lines.Add("remoteapplicationmode:i:0");
            lines.Add("alternate shell:s:");
            lines.Add("shell working directory:s:");
            lines.Add("gatewayhostname:s:");
            lines.Add("gatewayusagemethod:i:4");
            lines.Add("gatewaycredentialssource:i:4");
            lines.Add("gatewayprofileusagemethod:i:0");
            lines.Add("promptcredentialonce:i:0");
            lines.Add("gatewaybrokeringtype:i:0");
            lines.Add("use redirection server name:i:0");
            lines.Add("rdgiskdcproxy:i:0");
            lines.Add("kdcproxyname:s:");

            File.WriteAllLines(filePath, lines, Encoding.UTF8);
        }

        private void ParseRdpLine(string line, RdpConnection connection)
        {
            if (string.IsNullOrWhiteSpace(line))
                return;

            var match = Regex.Match(line, @"^(.+?):[ish]:(.*)$");
            if (!match.Success)
                return;

            var key = match.Groups[1].Value.ToLower();
            var value = match.Groups[2].Value;

            switch (key)
            {
                case "full address":
                    connection.FullAddress = value;
                    break;
                case "username":
                    connection.Username = value;
                    break;
                case "domain":
                    connection.Domain = value;
                    break;
                case "screen mode id":
                    if (int.TryParse(value, out int screenMode))
                        connection.ScreenModeId = screenMode;
                    break;
                case "use multimon":
                    connection.UseMultimon = value == "1";
                    break;
                case "selectedmonitors":
                    connection.SelectedMonitors = ParseSelectedMonitors(value);
                    break;
                case "desktopwidth":
                    if (int.TryParse(value, out int width))
                        connection.DesktopWidth = width;
                    break;
                case "desktopheight":
                    if (int.TryParse(value, out int height))
                        connection.DesktopHeight = height;
                    break;
                case "session bpp":
                    if (int.TryParse(value, out int colorDepth))
                        connection.ColorDepth = colorDepth;
                    break;
                // エクスペリエンス設定
                case "connection type":
                    if (int.TryParse(value, out int connectionType))
                        connection.ConnectionType = connectionType;
                    break;
                case "disable wallpaper":
                    connection.DesktopBackground = value != "1";
                    break;
                case "allow font smoothing":
                    connection.FontSmoothing = value == "1";
                    break;
                case "allow desktop composition":
                    connection.DesktopComposition = value == "1";
                    break;
                case "disable full window drag":
                    connection.ShowWindowContents = value != "1";
                    break;
                case "disable menu anims":
                    connection.MenuAnimations = value != "1";
                    break;
                case "disable themes":
                    connection.VisualStyles = value != "1";
                    break;
                case "bitmapcachepersistenable":
                    connection.BitmapCaching = value == "1";
                    break;
                case "autoreconnection enabled":
                    connection.AutoReconnect = value == "1";
                    break;
                // ローカルリソース設定
                case "audiomode":
                    if (int.TryParse(value, out int audioMode))
                        connection.AudioMode = audioMode;
                    break;
                case "audiocapturemode":
                    connection.AudioRecord = value == "1";
                    break;
                case "keyboardhook":
                    if (int.TryParse(value, out int keyboardMode))
                        connection.KeyboardMode = keyboardMode;
                    break;
                case "redirectprinters":
                    connection.RedirectPrinters = value == "1";
                    break;
                case "redirectcomports":
                    connection.RedirectPorts = value == "1";
                    break;
                case "redirectsmartcards":
                    connection.RedirectSmartCards = value == "1";
                    break;
                case "redirectclipboard":
                    connection.RedirectClipboard = value == "1";
                    break;
                case "redirectdrives":
                    connection.RedirectDrives = value == "1";
                    break;
                case "redirectposdevices":
                    connection.RedirectPnpDevices = value == "1";
                    break;
            }
        }

        private int ParseSelectedMonitors(string value)
        {
            // selectedmonitorsは "0,1,2" のような形式
            // ビットフラグに変換
            if (string.IsNullOrEmpty(value))
                return 0;

            int result = 0;
            var monitors = value.Split(',');
            foreach (var monitor in monitors)
            {
                if (int.TryParse(monitor.Trim(), out int monitorIndex))
                {
                    result |= (1 << monitorIndex);
                }
            }
            return result;
        }

        private string BuildSelectedMonitorsString(int selectedMonitors)
        {
            var monitors = new List<int>();
            for (int i = 0; i < 32; i++)
            {
                if ((selectedMonitors & (1 << i)) != 0)
                {
                    monitors.Add(i);
                }
            }
            return string.Join(",", monitors);
        }
    }
}