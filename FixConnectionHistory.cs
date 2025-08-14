using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace RemoteWakeConnect
{
    /// <summary>
    /// 接続履歴のデータを修正するユーティリティプログラム
    /// コンピュータ名とIPアドレスが混在している問題を修正
    /// </summary>
    public class FixConnectionHistory
    {
        // RdpConnectionクラスの簡易版
        public class ConnectionItem
        {
            public string Name { get; set; } = string.Empty;
            public string FullAddress { get; set; } = string.Empty;
            public string ComputerName { get; set; } = string.Empty;
            public string IpAddressValue { get; set; } = string.Empty;
            public int Port { get; set; } = 3389;
            public string Username { get; set; } = string.Empty;
            public string Domain { get; set; } = string.Empty;
            public string MacAddress { get; set; } = string.Empty;
            public int ScreenModeId { get; set; } = 2;
            public bool UseMultimon { get; set; } = false;
            public int SelectedMonitors { get; set; } = 0;
            public int DesktopWidth { get; set; } = 1920;
            public int DesktopHeight { get; set; } = 1080;
            public DateTime LastConnection { get; set; }
            public int SavedMonitorCount { get; set; } = 0;
            public List<int> SelectedMonitorIndices { get; set; } = new List<int>();
            public string MonitorConfigHash { get; set; } = string.Empty;
            
            // 互換性のためのプロパティ
            public string IpAddress { get; set; } = string.Empty;
        }

        public static void Main(string[] args)
        {
            Console.WriteLine("=== 接続履歴修正プログラム ===");
            Console.WriteLine();

            try
            {
                // 履歴ファイルのパスを取得
                string historyFilePath = Path.Combine(
                    AppContext.BaseDirectory,
                    "connection_history.yaml"
                );

                if (!File.Exists(historyFilePath))
                {
                    Console.WriteLine($"履歴ファイルが見つかりません: {historyFilePath}");
                    Console.WriteLine("\nEnterキーを押して終了...");
                    Console.ReadLine();
                    return;
                }

                Console.WriteLine($"履歴ファイルを読み込み中: {historyFilePath}");

                // バックアップを作成
                string backupPath = historyFilePath + $".backup_{DateTime.Now:yyyyMMdd_HHmmss}";
                File.Copy(historyFilePath, backupPath, true);
                Console.WriteLine($"バックアップを作成しました: {backupPath}");

                // YAMLデシリアライザーを設定
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .IgnoreUnmatchedProperties()
                    .Build();

                // 履歴データを読み込み
                string yamlContent = File.ReadAllText(historyFilePath);
                var history = deserializer.Deserialize<List<ConnectionItem>>(yamlContent) ?? new List<ConnectionItem>();

                Console.WriteLine($"\n{history.Count} 件の接続履歴を読み込みました。");
                Console.WriteLine("\n修正処理を開始...\n");

                int fixedCount = 0;
                foreach (var item in history)
                {
                    bool modified = false;
                    string originalInfo = $"[{item.FullAddress}]";

                    // ComputerNameとIpAddressValueが空で、IpAddressに値がある場合
                    if (string.IsNullOrEmpty(item.ComputerName) && 
                        string.IsNullOrEmpty(item.IpAddressValue) && 
                        !string.IsNullOrEmpty(item.IpAddress))
                    {
                        // IPアドレスかコンピュータ名かを判定
                        if (IsIpAddress(item.IpAddress))
                        {
                            item.IpAddressValue = item.IpAddress;
                            item.ComputerName = "";
                            Console.WriteLine($"  {originalInfo} → IP: {item.IpAddressValue}");
                        }
                        else
                        {
                            item.ComputerName = item.IpAddress;
                            item.IpAddressValue = "";
                            Console.WriteLine($"  {originalInfo} → コンピュータ名: {item.ComputerName}");
                        }
                        modified = true;
                    }
                    
                    // FullAddressから情報を抽出（必要に応じて）
                    if (!modified && string.IsNullOrEmpty(item.ComputerName) && 
                        string.IsNullOrEmpty(item.IpAddressValue) && 
                        !string.IsNullOrEmpty(item.FullAddress))
                    {
                        var parts = item.FullAddress.Split(':');
                        var hostPart = parts[0];
                        
                        if (IsIpAddress(hostPart))
                        {
                            item.IpAddressValue = hostPart;
                            item.ComputerName = "";
                            Console.WriteLine($"  {originalInfo} → IP: {item.IpAddressValue} (FullAddressから抽出)");
                        }
                        else
                        {
                            item.ComputerName = hostPart;
                            item.IpAddressValue = "";
                            Console.WriteLine($"  {originalInfo} → コンピュータ名: {item.ComputerName} (FullAddressから抽出)");
                        }
                        modified = true;
                    }

                    // IpAddressプロパティをクリア（後方互換性のため残すが空にする）
                    if (!string.IsNullOrEmpty(item.IpAddress))
                    {
                        item.IpAddress = "";
                        modified = true;
                    }

                    if (modified)
                    {
                        fixedCount++;
                    }
                }

                Console.WriteLine($"\n{fixedCount} 件のデータを修正しました。");

                // YAMLシリアライザーを設定
                var serializer = new SerializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .Build();

                // 修正したデータを保存
                string updatedYaml = serializer.Serialize(history);
                File.WriteAllText(historyFilePath, updatedYaml);

                Console.WriteLine($"\n履歴ファイルを更新しました: {historyFilePath}");
                Console.WriteLine("\n修正処理が完了しました！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nエラーが発生しました: {ex.Message}");
                Console.WriteLine($"詳細: {ex.StackTrace}");
            }

            Console.WriteLine("\nEnterキーを押して終了...");
            Console.ReadLine();
        }

        /// <summary>
        /// 文字列がIPアドレスかどうかを判定
        /// </summary>
        private static bool IsIpAddress(string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;

            // IPアドレスのパターン（簡易版）
            var ipPattern = @"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$";
            if (Regex.IsMatch(value, ipPattern))
            {
                // 各オクテットが0-255の範囲内かチェック
                var parts = value.Split('.');
                return parts.All(part => int.TryParse(part, out int num) && num >= 0 && num <= 255);
            }

            // IPv6の簡易チェック
            if (value.Contains(":"))
            {
                return IPAddress.TryParse(value, out _);
            }

            return false;
        }
    }
}