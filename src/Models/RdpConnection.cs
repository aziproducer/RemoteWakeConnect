using System;
using System.Collections.Generic;

namespace RemoteWakeConnect.Models
{
    public class RdpConnection
    {
        public string Name { get; set; } = string.Empty;
        public string FullAddress { get; set; } = string.Empty;
        public string ComputerName { get; set; } = string.Empty;  // コンピュータ名
        public string IpAddressValue { get; set; } = string.Empty;  // IPアドレス
        public int Port { get; set; } = 3389;
        public string Username { get; set; } = string.Empty;
        public string Domain { get; set; } = string.Empty;
        public string MacAddress { get; set; } = string.Empty;
        public string RdpFilePath { get; set; } = string.Empty; // 履歴用RDPファイルのパス
        public int ScreenModeId { get; set; } = 2; // 1=Windowed, 2=Fullscreen
        public bool UseMultimon { get; set; } = false;
        public int SelectedMonitors { get; set; } = 0;
        public int DesktopWidth { get; set; } = 1920;
        public int DesktopHeight { get; set; } = 1080;
        public int ColorDepth { get; set; } = 24; // 16, 24, 32 bits
        public DateTime LastConnection { get; set; }
        
        // モニター設定の保存
        public int SavedMonitorCount { get; set; } = 0;
        public List<int> SelectedMonitorIndices { get; set; } = new List<int>();
        public string MonitorConfigHash { get; set; } = string.Empty; // モニター構成のハッシュ値
        
        // エクスペリエンス設定
        public int ConnectionType { get; set; } = 0; // 0=自動検出, 1=LAN, 2=ブロードバンド, 3=モデム, 4=カスタム
        public bool DesktopBackground { get; set; } = true;
        public bool FontSmoothing { get; set; } = true;
        public bool DesktopComposition { get; set; } = true;
        public bool ShowWindowContents { get; set; } = true;
        public bool MenuAnimations { get; set; } = true;
        public bool VisualStyles { get; set; } = true;
        public bool BitmapCaching { get; set; } = true;
        public bool AutoReconnect { get; set; } = true;
        
        // ローカルリソース設定
        public int AudioMode { get; set; } = 0; // 0=ローカル, 1=リモート, 2=再生しない
        public bool AudioRecord { get; set; } = false;
        public int KeyboardMode { get; set; } = 0; // 0=ローカル, 1=リモート, 2=全画面時のみリモート
        public bool RedirectPrinters { get; set; } = true;
        public bool RedirectClipboard { get; set; } = true;
        public bool RedirectSmartCards { get; set; } = false;
        public bool RedirectPorts { get; set; } = false;
        public bool RedirectDrives { get; set; } = false;
        public bool RedirectPnpDevices { get; set; } = false;
        
        // OS情報のキャッシュ（セッション監視の高速化）
        public string CachedOsType { get; set; } = string.Empty; // "Workstation", "ServerWithRds", "ServerWithoutRds", "Unknown"
        public bool CachedIsRdsInstalled { get; set; } = false;
        public int CachedMaxSessions { get; set; } = 0;
        public DateTime CachedOsInfoTime { get; set; } = DateTime.MinValue;
        
        // 互換性のためのプロパティ（FullAddressから値を生成）
        public string IpAddress 
        { 
            get 
            {
                // 新しいIpAddressValueがあればそれを返す
                if (!string.IsNullOrEmpty(IpAddressValue))
                    return IpAddressValue;
                    
                // 互換性のため、FullAddressから抽出
                if (string.IsNullOrEmpty(FullAddress))
                    return string.Empty;
                
                var parts = FullAddress.Split(':');
                return parts[0];
            }
            set
            {
                IpAddressValue = value;
                // FullAddressも更新
                UpdateFullAddress();
            }
        }
        
        // FullAddressを更新するヘルパーメソッド
        public void UpdateFullAddress()
        {
            if (!string.IsNullOrEmpty(ComputerName))
            {
                FullAddress = Port != 3389 ? $"{ComputerName}:{Port}" : ComputerName;
            }
            else if (!string.IsNullOrEmpty(IpAddressValue))
            {
                FullAddress = Port != 3389 ? $"{IpAddressValue}:{Port}" : IpAddressValue;
            }
        }

        public RdpConnection Clone()
        {
            return new RdpConnection
            {
                Name = this.Name,
                FullAddress = this.FullAddress,
                ComputerName = this.ComputerName,
                IpAddressValue = this.IpAddressValue,
                Port = this.Port,
                Username = this.Username,
                Domain = this.Domain,
                MacAddress = this.MacAddress,
                RdpFilePath = this.RdpFilePath,
                ScreenModeId = this.ScreenModeId,
                UseMultimon = this.UseMultimon,
                SelectedMonitors = this.SelectedMonitors,
                DesktopWidth = this.DesktopWidth,
                DesktopHeight = this.DesktopHeight,
                ColorDepth = this.ColorDepth,
                LastConnection = this.LastConnection,
                SavedMonitorCount = this.SavedMonitorCount,
                SelectedMonitorIndices = new List<int>(this.SelectedMonitorIndices),
                MonitorConfigHash = this.MonitorConfigHash,
                
                // エクスペリエンス設定
                ConnectionType = this.ConnectionType,
                DesktopBackground = this.DesktopBackground,
                FontSmoothing = this.FontSmoothing,
                DesktopComposition = this.DesktopComposition,
                ShowWindowContents = this.ShowWindowContents,
                MenuAnimations = this.MenuAnimations,
                VisualStyles = this.VisualStyles,
                BitmapCaching = this.BitmapCaching,
                AutoReconnect = this.AutoReconnect,
                
                // ローカルリソース設定
                AudioMode = this.AudioMode,
                AudioRecord = this.AudioRecord,
                KeyboardMode = this.KeyboardMode,
                RedirectPrinters = this.RedirectPrinters,
                RedirectClipboard = this.RedirectClipboard,
                RedirectSmartCards = this.RedirectSmartCards,
                RedirectPorts = this.RedirectPorts,
                RedirectDrives = this.RedirectDrives,
                RedirectPnpDevices = this.RedirectPnpDevices,
                
                // OS情報のキャッシュ
                CachedOsType = this.CachedOsType,
                CachedIsRdsInstalled = this.CachedIsRdsInstalled,
                CachedMaxSessions = this.CachedMaxSessions,
                CachedOsInfoTime = this.CachedOsInfoTime
            };
        }
    }
}