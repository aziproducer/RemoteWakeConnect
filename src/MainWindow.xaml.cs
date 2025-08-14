using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Microsoft.Win32;
using RemoteWakeConnect.Models;
using RemoteWakeConnect.Services;

namespace RemoteWakeConnect
{
    public partial class MainWindow : Window
    {
        private readonly MonitorService _monitorService;
        private readonly WakeOnLanService _wakeOnLanService;
        private readonly RdpFileService _rdpFileService;
        private readonly RemoteDesktopService _remoteDesktopService;
        private readonly ConnectionHistoryService _historyService;
        private readonly NetworkService _networkService;
        private readonly MonitorConfigService _monitorConfigService;
        
        private List<MonitorInfo> _currentMonitors;
        private RdpConnection? _currentConnection;
        private List<Rectangle> _monitorRectangles;

        private static readonly string LogFile = System.IO.Path.Combine(
            AppContext.BaseDirectory,
            "debug.log"
        );

        public MainWindow()
        {
            try
            {
                LogDebug("MainWindow constructor started");
                
                LogDebug("Calling InitializeComponent");
                InitializeComponent();
                LogDebug("InitializeComponent completed");
                
                LogDebug("Creating services");
                _monitorService = new MonitorService();
                _wakeOnLanService = new WakeOnLanService();
                _rdpFileService = new RdpFileService();
                _remoteDesktopService = new RemoteDesktopService();
                _historyService = new ConnectionHistoryService();
                _networkService = new NetworkService();
                _monitorConfigService = new MonitorConfigService();
                LogDebug("Services created");
                
                _currentMonitors = new List<MonitorInfo>();
                _monitorRectangles = new List<Rectangle>();
                
                LogDebug("Calling Initialize");
                Initialize();
                LogDebug("Initialize completed");
            }
            catch (Exception ex)
            {
                LogError("MainWindow constructor failed", ex);
                throw;
            }
        }

        private static void LogDebug(string message)
        {
            try
            {
                File.AppendAllText(LogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [DEBUG] {message}\n");
            }
            catch { }
        }

        private static void LogError(string message, Exception ex)
        {
            try
            {
                var errorMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [ERROR] {message}\n";
                errorMessage += $"  Exception: {ex.GetType().Name}\n";
                errorMessage += $"  Message: {ex.Message}\n";
                errorMessage += $"  StackTrace:\n{ex.StackTrace}\n";
                
                if (ex.InnerException != null)
                {
                    errorMessage += $"  Inner Exception: {ex.InnerException.GetType().Name}\n";
                    errorMessage += $"  Inner Message: {ex.InnerException.Message}\n";
                }
                
                File.AppendAllText(LogFile, errorMessage);
            }
            catch { }
        }

        private void Initialize()
        {
            try
            {
                LogDebug("Initialize: Starting RefreshMonitorInfo");
                RefreshMonitorInfo();
                LogDebug("Initialize: RefreshMonitorInfo completed");
                
                LogDebug("Initialize: Loading connection history");
                ConnectionHistoryGrid.ItemsSource = _historyService.GetHistory();
                LogDebug("Initialize: Connection history loaded");
                
                // コマンドライン引数からの起動処理
                ProcessStartupArgs();
            }
            catch (Exception ex)
            {
                LogError("Initialize failed", ex);
            }
        }

        private void BrowseRdpButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "RDPファイル (*.rdp)|*.rdp|すべてのファイル (*.*)|*.*",
                Title = "RDPファイルを選択"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                RdpFilePathTextBox.Text = openFileDialog.FileName;
                LoadRdpFile(openFileDialog.FileName);
            }
        }

        private void RdpFilePathTextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void RdpFilePathTextBox_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && files[0].EndsWith(".rdp", StringComparison.OrdinalIgnoreCase))
                {
                    RdpFilePathTextBox.Text = files[0];
                    LoadRdpFile(files[0]);
                }
            }
        }

        private async void LoadRdpFile(string filePath)
        {
            try
            {
                _currentConnection = _rdpFileService.LoadRdpFile(filePath);
                
                // FullAddressからコンピュータ名またはIPとポートを分離
                var fullAddr = _currentConnection.FullAddress;
                if (!string.IsNullOrEmpty(fullAddr))
                {
                    var parts = fullAddr.Split(':');
                    var hostPart = parts[0];
                    
                    // ポート番号を取得（デフォルトは3389）
                    if (parts.Length > 1 && int.TryParse(parts[1], out int port))
                    {
                        _currentConnection.Port = port;
                        PortTextBox.Text = port.ToString();
                    }
                    else
                    {
                        PortTextBox.Text = "3389";
                    }
                    
                    // IPアドレスかコンピュータ名か判定
                    if (System.Net.IPAddress.TryParse(hostPart, out _))
                    {
                        IpAddressTextBox.Text = hostPart;
                        ComputerNameTextBox.Text = "";
                        DirectAddressTextBox.Text = hostPart;
                        _currentConnection.IpAddressValue = hostPart;
                    }
                    else
                    {
                        ComputerNameTextBox.Text = hostPart;
                        IpAddressTextBox.Text = "";
                        DirectAddressTextBox.Text = hostPart;
                        _currentConnection.ComputerName = hostPart;
                        // コンピュータ名からIPを解決を試みる
                        _ = ResolveHostNameAsync(hostPart);
                    }
                }
                
                // 履歴から接続情報を検索
                var historicalConnection = _historyService.FindByAddress(_currentConnection.FullAddress);
                if (historicalConnection != null)
                {
                    // MACアドレスを復元
                    if (!string.IsNullOrEmpty(historicalConnection.MacAddress))
                    {
                        MacAddressTextBox.Text = historicalConnection.MacAddress;
                        _currentConnection.MacAddress = historicalConnection.MacAddress;
                    }
                    
                    // モニター設定を復元
                    await RestoreMonitorSettingsAsync(historicalConnection);
                    
                    StatusText.Text = $"RDPファイルを読み込みました: {System.IO.Path.GetFileName(filePath)} (設定を履歴から復元)";
                }
                else
                {
                    // ネットワークからMACアドレスを取得を試みる
                    StatusText.Text = "MACアドレスを検索中...";
                    
                    string? macAddress = null;
                    
                    // IPアドレスまたはコンピュータ名からMACアドレスを取得
                    var targetHost = !string.IsNullOrEmpty(IpAddressTextBox.Text) ? IpAddressTextBox.Text : ComputerNameTextBox.Text;
                    if (!string.IsNullOrEmpty(targetHost))
                    {
                        macAddress = await _networkService.GetMacAddressAsync(targetHost);
                    }
                    
                    // 同時にPCの状態も確認
                    _ = CheckHostStatusAsync(targetHost);
                    
                    // nbtstatでも試してみる
                    if (string.IsNullOrEmpty(macAddress) && !string.IsNullOrEmpty(targetHost))
                    {
                        macAddress = await _networkService.GetMacFromNbtstatAsync(targetHost);
                    }
                    
                    if (!string.IsNullOrEmpty(macAddress))
                    {
                        MacAddressTextBox.Text = macAddress;
                        _currentConnection.MacAddress = macAddress;
                        StatusText.Text = $"RDPファイルを読み込みました: {System.IO.Path.GetFileName(filePath)} (MACアドレスを自動取得)";
                    }
                    else if (!string.IsNullOrEmpty(_currentConnection.MacAddress))
                    {
                        MacAddressTextBox.Text = _currentConnection.MacAddress;
                        StatusText.Text = $"RDPファイルを読み込みました: {System.IO.Path.GetFileName(filePath)}";
                    }
                    else
                    {
                        StatusText.Text = $"RDPファイルを読み込みました: {System.IO.Path.GetFileName(filePath)} (MACアドレスは取得できませんでした)";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"RDPファイルの読み込みに失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RefreshMonitorsButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshMonitorInfo();
        }

        private void RefreshMonitorInfo()
        {
            try
            {
                LogDebug("RefreshMonitorInfo: Starting GetMonitors");
                _currentMonitors = _monitorService.GetMonitors();
                LogDebug($"RefreshMonitorInfo: Got {_currentMonitors.Count} monitors");
                
                var monitorText = $"検出されたモニター: {_currentMonitors.Count}台\n";
                monitorText += "※チェックボックスとレイアウト図の番号が対応しています";
                
                LogDebug("RefreshMonitorInfo: Setting CurrentMonitorConfigText");
                if (CurrentMonitorConfigText != null)
                {
                    CurrentMonitorConfigText.Text = monitorText;
                    LogDebug("RefreshMonitorInfo: CurrentMonitorConfigText set");
                }
                else
                {
                    LogDebug("RefreshMonitorInfo: CurrentMonitorConfigText is null");
                }
                
                // モニターチェックボックスリストを更新
                LogDebug("RefreshMonitorInfo: Updating MonitorCheckBoxList");
                if (MonitorCheckBoxList != null)
                {
                    MonitorCheckBoxList.ItemsSource = _currentMonitors;
                    LogDebug("RefreshMonitorInfo: MonitorCheckBoxList updated");
                }
                
                LogDebug("RefreshMonitorInfo: Calling UpdateMonitorLayout");
                UpdateMonitorLayout();
                LogDebug("RefreshMonitorInfo: UpdateMonitorLayout completed");
            }
            catch (Exception ex)
            {
                LogError("RefreshMonitorInfo failed", ex);
            }
        }

        private void MonitorCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            UpdateMonitorLayout();
        }

        private void MonitorCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            UpdateMonitorLayout();
        }

        private void SelectAllMonitorsButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var monitor in _currentMonitors)
            {
                monitor.IsSelected = true;
            }
            MonitorCheckBoxList.Items.Refresh();
            UpdateMonitorLayout();
        }

        private void DeselectAllMonitorsButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var monitor in _currentMonitors)
            {
                monitor.IsSelected = false;
            }
            MonitorCheckBoxList.Items.Refresh();
            UpdateMonitorLayout();
        }

        private void MonitorLayoutCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateMonitorLayout();
        }

        private void UpdateMonitorLayout()
        {
            MonitorLayoutCanvas.Children.Clear();
            _monitorRectangles.Clear();
            
            if (_currentMonitors.Count == 0)
                return;
            
            // Canvasのサイズがまだ決まっていない場合はスキップ
            if (MonitorLayoutCanvas.ActualWidth <= 0 || MonitorLayoutCanvas.ActualHeight <= 0)
                return;
            
            var virtualBounds = _monitorService.GetVirtualScreenBounds();
            
            // 仮想スクリーンのサイズが異常な場合はスキップ
            if (virtualBounds.Width <= 0 || virtualBounds.Height <= 0)
                return;
            
            double scale = Math.Min(
                (MonitorLayoutCanvas.ActualWidth - 20) / virtualBounds.Width,
                (MonitorLayoutCanvas.ActualHeight - 20) / virtualBounds.Height
            );
            
            for (int i = 0; i < _currentMonitors.Count; i++)
            {
                var monitor = _currentMonitors[i];
                bool isSelected = monitor.IsSelected;
                
                // モニターのサイズを確認
                double rectWidth = Math.Max(1, monitor.Width * scale);
                double rectHeight = Math.Max(1, monitor.Height * scale);
                
                var rect = new Rectangle
                {
                    Width = rectWidth,
                    Height = rectHeight,
                    Fill = isSelected ? Brushes.LightBlue : Brushes.LightGray,
                    Stroke = monitor.IsPrimary ? Brushes.Red : Brushes.DarkGray,
                    StrokeThickness = monitor.IsPrimary ? 2 : 1,
                    Tag = monitor
                };
                
                Canvas.SetLeft(rect, (monitor.X - virtualBounds.X) * scale + 10);
                Canvas.SetTop(rect, (monitor.Y - virtualBounds.Y) * scale + 10);
                
                rect.MouseLeftButtonDown += MonitorRectangle_MouseLeftButtonDown;
                
                MonitorLayoutCanvas.Children.Add(rect);
                _monitorRectangles.Add(rect);
                
                // モニター番号を大きく表示
                var numberBorder = new Border
                {
                    Background = monitor.IsPrimary ? Brushes.OrangeRed : Brushes.DarkBlue,
                    CornerRadius = new CornerRadius(3),
                    Width = 25,
                    Height = 25
                };
                
                var label = new TextBlock
                {
                    Text = monitor.Index.ToString(),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                
                numberBorder.Child = label;
                
                Canvas.SetLeft(numberBorder, (monitor.X - virtualBounds.X) * scale + 10);
                Canvas.SetTop(numberBorder, (monitor.Y - virtualBounds.Y) * scale + 10);
                Canvas.SetZIndex(numberBorder, 10);
                
                // 解像度情報を表示
                var resolutionLabel = new TextBlock
                {
                    Text = monitor.Resolution,
                    FontSize = 10,
                    Foreground = Brushes.DarkGray
                };
                
                Canvas.SetLeft(resolutionLabel, (monitor.X - virtualBounds.X) * scale + 10 + (rect.Width / 2) - 30);
                Canvas.SetTop(resolutionLabel, (monitor.Y - virtualBounds.Y) * scale + 10 + (rect.Height / 2) - 10);
                
                MonitorLayoutCanvas.Children.Add(numberBorder);
                MonitorLayoutCanvas.Children.Add(resolutionLabel);
            }
        }

        private void MonitorRectangle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var rect = sender as Rectangle;
            if (rect == null) return;
            
            var monitor = rect.Tag as MonitorInfo;
            if (monitor == null) return;
            
            monitor.IsSelected = !monitor.IsSelected;
            rect.Fill = monitor.IsSelected ? Brushes.LightBlue : Brushes.LightGray;
            
            // チェックボックスリストも更新
            MonitorCheckBoxList.Items.Refresh();
        }

        private List<MonitorInfo> GetSelectedMonitors()
        {
            return _currentMonitors.Where(m => m.IsSelected).ToList();
        }

        private async void WakeButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(MacAddressTextBox.Text))
            {
                MessageBox.Show("MACアドレスを入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            try
            {
                WakeButton.IsEnabled = false;
                StatusText.Text = "Wake On LANパケットを送信中...";
                
                var targetHost = !string.IsNullOrEmpty(IpAddressTextBox.Text) ? IpAddressTextBox.Text : ComputerNameTextBox.Text;
                await _wakeOnLanService.SendMagicPacketAsync(MacAddressTextBox.Text, targetHost);
                StatusText.Text = "Wake On LANパケットを送信しました。";
                
                if (!string.IsNullOrWhiteSpace(targetHost))
                {
                    StatusText.Text += " 起動確認中...";
                    
                    // 複数回チェック（最大30秒）
                    bool isOnline = false;
                    for (int i = 0; i < 6; i++)
                    {
                        await Task.Delay(5000);
                        isOnline = await _wakeOnLanService.PingHostAsync(targetHost);
                        if (isOnline)
                        {
                            UpdateStatusIndicator(true, "PCは起動中です");
                            StatusText.Text = "PCが起動しました。";
                            break;
                        }
                        StatusText.Text = $"起動確認中... ({(i + 1) * 5}秒経過)";
                    }
                    
                    if (!isOnline)
                    {
                        UpdateStatusIndicator(false, "PCはオフラインまたは応答していません");
                        StatusText.Text = "PCの起動を確認できませんでした。";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wake On LANの送信に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "エラーが発生しました。";
            }
            finally
            {
                WakeButton.IsEnabled = true;
            }
        }

        private async void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ConnectButton.IsEnabled = false;
                StatusText.Text = "リモートデスクトップに接続中...";
                
                if (_currentConnection == null)
                {
                    _currentConnection = new RdpConnection();
                }
                
                // RDPファイルからの設定がある場合
                if (!string.IsNullOrEmpty(ComputerNameTextBox.Text) || !string.IsNullOrEmpty(IpAddressTextBox.Text))
                {
                    _currentConnection.ComputerName = ComputerNameTextBox.Text;
                    _currentConnection.IpAddressValue = IpAddressTextBox.Text;
                    _currentConnection.MacAddress = MacAddressTextBox.Text;
                }
                // 直接入力からの設定
                else if (!string.IsNullOrEmpty(DirectAddressTextBox.Text))
                {
                    var directAddress = DirectAddressTextBox.Text;
                    
                    // ポート番号を取得
                    if (!string.IsNullOrEmpty(PortTextBox.Text) && int.TryParse(PortTextBox.Text, out int port))
                    {
                        _currentConnection.Port = port;
                    }
                    else
                    {
                        _currentConnection.Port = 3389;
                    }
                    
                    // アドレスがIPかコンピュータ名か判定
                    if (System.Net.IPAddress.TryParse(directAddress, out _))
                    {
                        _currentConnection.IpAddressValue = directAddress;
                        _currentConnection.ComputerName = "";
                    }
                    else
                    {
                        _currentConnection.ComputerName = directAddress;
                        _currentConnection.IpAddressValue = "";
                    }
                    
                    // ユーザー名を設定
                    if (!string.IsNullOrEmpty(UsernameTextBox.Text))
                    {
                        _currentConnection.Username = UsernameTextBox.Text;
                    }
                    
                    _currentConnection.MacAddress = MacAddressTextBox.Text;
                }
                else
                {
                    MessageBox.Show("接続先を指定してください。\nRDPファイルを選択するか、直接接続設定にアドレスを入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // FullAddressを更新（ポート番号も考慮）
                _currentConnection.UpdateFullAddress();
                
                // モニター設定を反映
                var selectedMonitors = GetSelectedMonitors();
                if (selectedMonitors.Count > 0)
                {
                    _currentConnection.UseMultimon = selectedMonitors.Count > 1;
                    _currentConnection.SelectedMonitors = _monitorService.BuildSelectedMonitorsFlag(selectedMonitors);
                    
                    // モニター設定を保存
                    _currentConnection.SavedMonitorCount = _currentMonitors.Count;
                    _currentConnection.SelectedMonitorIndices = _monitorConfigService.GetSelectedMonitorIndices(_currentMonitors);
                    _currentConnection.MonitorConfigHash = _monitorConfigService.GenerateMonitorConfigHash(_currentMonitors);
                }
                else
                {
                    // モニターが選択されていない場合はプライマリモニターを使用
                    var primaryMonitor = _currentMonitors.FirstOrDefault(m => m.IsPrimary) ?? _currentMonitors.FirstOrDefault();
                    if (primaryMonitor != null)
                    {
                        _currentConnection.UseMultimon = false;
                        _currentConnection.SelectedMonitors = _monitorService.BuildSelectedMonitorsFlag(new List<MonitorInfo> { primaryMonitor });
                    }
                }
                
                await _remoteDesktopService.ConnectAsync(_currentConnection);
                StatusText.Text = "リモートデスクトップ接続を開始しました。";
                
                // 接続履歴に追加（モニター設定を含む）
                _historyService.AddConnection(_currentConnection);
                
                // ジャンプリストを更新
                ((App)Application.Current).UpdateJumpList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"リモートデスクトップ接続に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "エラーが発生しました。";
            }
            finally
            {
                ConnectButton.IsEnabled = true;
            }
        }

        // 新しいイベントハンドラー
        private void HistorySearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = HistorySearchTextBox.Text?.ToLower() ?? "";
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // 検索文字列が空の場合は全て表示
                ConnectionHistoryGrid.ItemsSource = _historyService.GetHistory();
            }
            else
            {
                // フィルタリング
                var filtered = _historyService.GetHistory()
                    .Where(c => 
                        (c.FullAddress?.ToLower().Contains(searchText) ?? false) ||
                        (c.ComputerName?.ToLower().Contains(searchText) ?? false) ||
                        (c.IpAddressValue?.ToLower().Contains(searchText) ?? false) ||
                        (c.MacAddress?.ToLower().Contains(searchText) ?? false))
                    .ToList();
                ConnectionHistoryGrid.ItemsSource = filtered;
            }
        }

        private void ClearHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("接続履歴をすべて削除しますか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                _historyService.ClearHistory();
                StatusText.Text = "接続履歴をクリアしました。";
            }
        }

        private void ConnectionHistoryGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ConnectionHistoryGrid.SelectedItem != null)
            {
                UseSelectedHistory();
            }
        }

        private void UseHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button?.Tag != null)
            {
                UseSelectedHistory();
            }
        }
        
        private async Task ResolveHostNameAsync(string hostname)
        {
            try
            {
                var hostEntry = await Task.Run(() => System.Net.Dns.GetHostEntry(hostname));
                if (hostEntry.AddressList.Length > 0)
                {
                    // IPv4アドレスを優先
                    var ipv4 = hostEntry.AddressList.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                    if (ipv4 != null)
                    {
                        IpAddressTextBox.Text = ipv4.ToString();
                    }
                    else if (hostEntry.AddressList.Length > 0)
                    {
                        IpAddressTextBox.Text = hostEntry.AddressList[0].ToString();
                    }
                }
            }
            catch
            {
                // 解決できなかった場合は何もしない
            }
        }

        private async void UseSelectedHistory()
        {
            var selectedItem = ConnectionHistoryGrid.SelectedItem as RdpConnection;
            if (selectedItem != null)
            {
                // 接続設定を読み込み
                _currentConnection = selectedItem.Clone();
                
                // UIに反映
                ComputerNameTextBox.Text = _currentConnection.ComputerName ?? "";
                IpAddressTextBox.Text = _currentConnection.IpAddressValue ?? "";
                MacAddressTextBox.Text = _currentConnection.MacAddress ?? "";
                RdpFilePathTextBox.Text = _currentConnection.Name ?? "";
                
                // 直接接続設定にも反映
                if (!string.IsNullOrEmpty(_currentConnection.ComputerName))
                {
                    DirectAddressTextBox.Text = _currentConnection.ComputerName;
                }
                else if (!string.IsNullOrEmpty(_currentConnection.IpAddressValue))
                {
                    DirectAddressTextBox.Text = _currentConnection.IpAddressValue;
                }
                
                PortTextBox.Text = _currentConnection.Port != 3389 ? _currentConnection.Port.ToString() : "3389";
                UsernameTextBox.Text = _currentConnection.Username ?? "";
                
                // モニター設定を復元
                await RestoreMonitorSettingsAsync(_currentConnection);
                
                // 履歴から読み込まれたことを表示
                StatusText.Text = "履歴から接続設定を読み込みました。";
            }
        }

        private async void CheckStatusButton_Click(object sender, RoutedEventArgs e)
        {
            var targetHost = !string.IsNullOrEmpty(IpAddressTextBox.Text) ? IpAddressTextBox.Text : ComputerNameTextBox.Text;
            await CheckHostStatusAsync(targetHost);
        }

        private async Task CheckHostStatusAsync(string hostNameOrAddress)
        {
            if (string.IsNullOrWhiteSpace(hostNameOrAddress))
            {
                UpdateStatusIndicator(null, "アドレスが入力されていません");
                return;
            }

            try
            {
                StatusText.Text = "PCの状態を確認中...";
                UpdateStatusIndicator(null, "確認中...");
                
                bool isOnline = await _wakeOnLanService.PingHostAsync(hostNameOrAddress);
                
                if (isOnline)
                {
                    UpdateStatusIndicator(true, "PCは起動中です");
                    StatusText.Text = $"{hostNameOrAddress} は起動中です";
                }
                else
                {
                    UpdateStatusIndicator(false, "PCはオフラインまたは応答していません");
                    StatusText.Text = $"{hostNameOrAddress} はオフラインまたは応答していません";
                }
            }
            catch (Exception ex)
            {
                UpdateStatusIndicator(null, $"エラー: {ex.Message}");
                StatusText.Text = $"状態確認エラー: {ex.Message}";
            }
        }

        private void UpdateStatusIndicator(bool? isOnline, string tooltip)
        {
            if (StatusIndicator == null) return;
            
            StatusIndicator.ToolTip = tooltip;
            
            if (isOnline == null)
            {
                StatusIndicator.Background = System.Windows.Media.Brushes.Gray;
            }
            else if (isOnline.Value)
            {
                StatusIndicator.Background = System.Windows.Media.Brushes.LimeGreen;
            }
            else
            {
                StatusIndicator.Background = System.Windows.Media.Brushes.OrangeRed;
            }
        }

        private Task RestoreMonitorSettingsAsync(RdpConnection connection)
        {
            if (connection == null || connection.SelectedMonitorIndices == null || connection.SelectedMonitorIndices.Count == 0)
                return Task.CompletedTask;

            // 現在のモニター構成を取得
            if (_currentMonitors == null || _currentMonitors.Count == 0)
            {
                RefreshMonitorInfo();
            }
            
            // _currentMonitorsがまだnullの場合は処理を中止
            if (_currentMonitors == null || _currentMonitors.Count == 0)
                return Task.CompletedTask;

            // モニター構成が変更されているかチェック
            if (!string.IsNullOrEmpty(connection.MonitorConfigHash))
            {
                bool hasChanged = _monitorConfigService.HasMonitorConfigChanged(
                    connection.MonitorConfigHash, 
                    _currentMonitors
                );

                if (hasChanged)
                {
                    string changeDescription = _monitorConfigService.GetMonitorConfigChangeDescription(
                        connection.SavedMonitorCount,
                        _currentMonitors.Count
                    );

                    var result = MessageBox.Show(
                        $"前回の接続時からモニター構成が変更されています。\n\n" +
                        $"{changeDescription}\n\n" +
                        $"前回の設定を復元しますか？\n" +
                        $"「はい」: 前回の設定を復元\n" +
                        $"「いいえ」: 現在の設定を維持",
                        "モニター構成の変更",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning
                    );

                    if (result == MessageBoxResult.No)
                    {
                        return Task.CompletedTask;
                    }
                }
            }

            // モニター選択を復元
            _monitorConfigService.RestoreMonitorSelection(_currentMonitors, connection.SelectedMonitorIndices);
            
            // UIを更新
            if (MonitorCheckBoxList != null)
            {
                MonitorCheckBoxList.Items.Refresh();
            }
            UpdateMonitorLayout();
            
            StatusText.Text = "モニター設定を復元しました。";
            
            return Task.CompletedTask;
        }

        private async void ProcessStartupArgs()
        {
            try
            {
                var app = Application.Current;
                if (app.Properties.Contains("ConnectAddress"))
                {
                    string? address = app.Properties["ConnectAddress"]?.ToString();
                    if (string.IsNullOrEmpty(address))
                        return;
                        
                    bool useWol = app.Properties.Contains("UseWakeOnLan") && app.Properties["UseWakeOnLan"] is bool b && b;
                    string macAddress = app.Properties.Contains("MacAddress") ? app.Properties["MacAddress"]?.ToString() ?? "" : "";

                    // 履歴から接続情報を取得
                    var connection = _historyService.FindByAddress(address);
                    if (connection != null)
                    {
                        _currentConnection = connection.Clone();
                        
                        // UIに完全に反映
                        IpAddressTextBox.Text = _currentConnection.IpAddress;
                        MacAddressTextBox.Text = !string.IsNullOrEmpty(macAddress) ? macAddress : _currentConnection.MacAddress ?? "";
                        RdpFilePathTextBox.Text = _currentConnection.Name ?? "";
                        
                        // メインウィンドウにフォーカス
                        
                        // モニター設定を復元
                        await RestoreMonitorSettingsAsync(_currentConnection);
                        
                        // PC状態を確認
                        _ = CheckHostStatusAsync(_currentConnection.IpAddress);

                        if (useWol && !string.IsNullOrEmpty(MacAddressTextBox.Text))
                        {
                            // Wake On LANを送信
                            StatusText.Text = "Wake On LANパケットを送信中...";
                            var targetHost = !string.IsNullOrEmpty(IpAddressTextBox.Text) ? IpAddressTextBox.Text : ComputerNameTextBox.Text;
                            await _wakeOnLanService.SendMagicPacketAsync(MacAddressTextBox.Text, targetHost);
                            
                            // 起動を待つ
                            StatusText.Text = "PCの起動を待っています...";
                            bool isOnline = false;
                            for (int i = 0; i < 12; i++) // 最大60秒待つ
                            {
                                await Task.Delay(5000);
                                isOnline = await _wakeOnLanService.PingHostAsync(targetHost);
                                if (isOnline)
                                {
                                    UpdateStatusIndicator(true, "PCは起動中です");
                                    break;
                                }
                                StatusText.Text = $"起動確認中... ({(i + 1) * 5}秒経過)";
                            }
                            
                            if (!isOnline)
                            {
                                MessageBox.Show(
                                    "PCの起動を確認できませんでした。\n接続を続行しますか？",
                                    "起動確認",
                                    MessageBoxButton.YesNo,
                                    MessageBoxImage.Question
                                );
                                if (MessageBox.Show(
                                    "PCの起動を確認できませんでした。\n接続を続行しますか？",
                                    "起動確認",
                                    MessageBoxButton.YesNo,
                                    MessageBoxImage.Question) == MessageBoxResult.No)
                                {
                                    return;
                                }
                            }
                        }

                        // 接続を実行
                        StatusText.Text = "リモートデスクトップに接続中...";
                        
                        // 少し待機してUIが更新されるのを確認できるようにする
                        await Task.Delay(500);
                        
                        await _remoteDesktopService.ConnectAsync(_currentConnection);
                        
                        // 履歴を更新（モニター設定も含む）
                        _currentConnection.SavedMonitorCount = _currentMonitors?.Count ?? 0;
                        _currentConnection.SelectedMonitorIndices = _monitorConfigService.GetSelectedMonitorIndices(_currentMonitors ?? new List<MonitorInfo>());
                        _currentConnection.MonitorConfigHash = _monitorConfigService.GenerateMonitorConfigHash(_currentMonitors ?? new List<MonitorInfo>());
                        _historyService.AddConnection(_currentConnection);
                        
                        // ジャンプリストを更新
                        ((App)Application.Current).UpdateJumpList();
                        
                        StatusText.Text = "リモートデスクトップ接続を開始しました。";
                    }
                    else
                    {
                        MessageBox.Show(
                            $"指定されたアドレス '{address}' の接続履歴が見つかりません。",
                            "エラー",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("ProcessStartupArgs failed", ex);
                MessageBox.Show(
                    $"自動接続に失敗しました: {ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }
    }
}