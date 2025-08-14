using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using RemoteWakeConnect.Models;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace RemoteWakeConnect.Services
{
    public class ConnectionHistoryService
    {
        private readonly string _historyFilePath;
        private readonly ISerializer _serializer;
        private readonly IDeserializer _deserializer;
        private ObservableCollection<RdpConnection> _history;
        private const int MaxHistoryItems = 50;

        public ConnectionHistoryService()
        {
            // YAMLファイルのパスを設定（実行ファイルと同じフォルダ）
            _historyFilePath = Path.Combine(
                AppContext.BaseDirectory,
                "connection_history.yaml"
            );

            // YAMLシリアライザーの設定
            _serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();

            _deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();

            _history = new ObservableCollection<RdpConnection>();
            LoadHistory();
        }

        public ObservableCollection<RdpConnection> GetHistory()
        {
            return _history;
        }

        public void AddConnection(RdpConnection connection)
        {
            // 接続時刻を更新
            connection.LastConnection = DateTime.Now;
            
            // ComputerNameとIpAddressValueが正しく設定されていることを確認
            // IpAddressValueには必ずIPアドレスのみを格納
            if (!string.IsNullOrEmpty(connection.IpAddressValue))
            {
                // IPアドレスでない場合はComputerNameに移動
                if (!System.Net.IPAddress.TryParse(connection.IpAddressValue, out _))
                {
                    if (string.IsNullOrEmpty(connection.ComputerName))
                    {
                        connection.ComputerName = connection.IpAddressValue;
                    }
                    connection.IpAddressValue = "";
                }
            }

            // 同じアドレスの既存エントリを検索
            var existingIndex = _history.ToList().FindIndex(c => 
                c.FullAddress == connection.FullAddress);

            if (existingIndex >= 0)
            {
                // 既存エントリを更新（最新情報で上書き）
                _history[existingIndex] = connection.Clone();
                // リストの先頭に移動
                var item = _history[existingIndex];
                _history.RemoveAt(existingIndex);
                _history.Insert(0, item);
            }
            else
            {
                // 新規エントリとして追加
                _history.Insert(0, connection.Clone());
            }

            // 履歴の最大数を制限
            while (_history.Count > MaxHistoryItems)
            {
                _history.RemoveAt(_history.Count - 1);
            }

            SaveHistory();
        }

        public void UpdateConnection(string address, string macAddress, string ipAddress)
        {
            var connection = _history.FirstOrDefault(c => c.FullAddress == address);
            if (connection != null)
            {
                // MACアドレスとIPアドレスを更新
                if (!string.IsNullOrEmpty(macAddress))
                    connection.MacAddress = macAddress;
                
                if (!string.IsNullOrEmpty(ipAddress))
                {
                    // IPアドレスがFullAddressに含まれていない場合は更新
                    if (!connection.FullAddress.Contains(ipAddress))
                    {
                        connection.FullAddress = ipAddress;
                    }
                }
                
                SaveHistory();
            }
        }

        public void RemoveConnection(RdpConnection connection)
        {
            _history.Remove(connection);
            SaveHistory();
        }

        public void ClearHistory()
        {
            _history.Clear();
            SaveHistory();
        }

        public RdpConnection? FindByAddress(string address)
        {
            return _history.FirstOrDefault(c => 
                c.FullAddress == address || c.IpAddress == address);
        }

        private void LoadHistory()
        {
            try
            {
                if (File.Exists(_historyFilePath))
                {
                    var yaml = File.ReadAllText(_historyFilePath);
                    if (!string.IsNullOrWhiteSpace(yaml))
                    {
                        var connections = _deserializer.Deserialize<List<RdpConnection>>(yaml);
                        if (connections != null)
                        {
                            _history.Clear();
                            foreach (var conn in connections)
                            {
                                // 読み込み時にもデータを修正
                                FixConnectionData(conn);
                                _history.Add(conn);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // エラーが発生しても続行（新規ファイルとして扱う）
                System.Diagnostics.Debug.WriteLine($"履歴の読み込みエラー: {ex.Message}");
            }
        }

        private void SaveHistory()
        {
            try
            {
                // 保存前にデータを修正
                foreach (var conn in _history)
                {
                    FixConnectionData(conn);
                }
                
                // YAMLとして保存（実行ファイルと同じフォルダなのでディレクトリ作成は不要）
                var yaml = _serializer.Serialize(_history.ToList());
                File.WriteAllText(_historyFilePath, yaml);
            }
            catch (Exception ex)
            {
                // 保存エラーは警告のみ
                System.Diagnostics.Debug.WriteLine($"履歴の保存エラー: {ex.Message}");
            }
        }
        
        private void FixConnectionData(RdpConnection connection)
        {
            // IpAddressValueに文字列（コンピュータ名）が入っている場合の修正
            if (!string.IsNullOrEmpty(connection.IpAddressValue))
            {
                // IPアドレスでない場合はComputerNameに移動
                if (!System.Net.IPAddress.TryParse(connection.IpAddressValue, out _))
                {
                    if (string.IsNullOrEmpty(connection.ComputerName))
                    {
                        connection.ComputerName = connection.IpAddressValue;
                    }
                    connection.IpAddressValue = "";
                }
            }
            
            // 古いIpAddressプロパティからの移行
            if (!string.IsNullOrEmpty(connection.IpAddress))
            {
                // IPアドレスかどうかチェック
                if (System.Net.IPAddress.TryParse(connection.IpAddress, out _))
                {
                    if (string.IsNullOrEmpty(connection.IpAddressValue))
                    {
                        connection.IpAddressValue = connection.IpAddress;
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(connection.ComputerName))
                    {
                        connection.ComputerName = connection.IpAddress;
                    }
                }
            }
            
            // FullAddressから情報を抽出（ComputerNameとIpAddressValueが両方空の場合）
            if (string.IsNullOrEmpty(connection.ComputerName) && 
                string.IsNullOrEmpty(connection.IpAddressValue) && 
                !string.IsNullOrEmpty(connection.FullAddress))
            {
                var parts = connection.FullAddress.Split(':');
                var hostPart = parts[0];
                
                if (System.Net.IPAddress.TryParse(hostPart, out _))
                {
                    connection.IpAddressValue = hostPart;
                }
                else
                {
                    connection.ComputerName = hostPart;
                }
            }
        }

        public string GetHistoryFilePath()
        {
            return _historyFilePath;
        }
    }
}