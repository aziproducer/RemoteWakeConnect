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
using RemoteWakeConnect.Dialogs;
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
        private readonly SessionMonitorService _sessionMonitorService;
        
        private List<MonitorInfo> _currentMonitors;
        private RdpConnection? _currentConnection;
        private List<Rectangle> _monitorRectangles;
        private SessionCheckResult? _lastSessionCheckResult;
        private System.Threading.Timer? _sessionCheckTimer;
        private bool _isCheckingSession = false;
        private DateTime _lastSessionCheckTime = DateTime.MinValue;  // 最後にセッション確認が成功した時刻
        private readonly TimeSpan _sessionCheckValidDuration = TimeSpan.FromSeconds(10);  // セッション確認結果の有効期間
        
        // 標準的なRDP解像度リスト（プライマリモニターの解像度も含める）
        private List<(int Width, int Height, string DisplayName)> _availableResolutions = new();

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
                _sessionMonitorService = new SessionMonitorService();
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
                
                LogDebug("Initialize: Initializing resolution settings");
                InitializeResolutionSettings();
                LogDebug("Initialize: Resolution settings initialized");
                
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
                // セッション状態をリセット
                if (SessionStatusText != null)
                {
                    SessionStatusText.Text = "未確認";
                    SessionStatusText.Foreground = new SolidColorBrush(Colors.Gray);
                }
                
                _currentConnection = _rdpFileService.LoadRdpFile(filePath);
                
                // 外部からのRDPファイルの場合、rdp_filesフォルダにコピー
                var rdpFilesFolder = System.IO.Path.Combine(AppContext.BaseDirectory, "rdp_files");
                if (!filePath.StartsWith(rdpFilesFolder, StringComparison.OrdinalIgnoreCase))
                {
                    // rdp_filesフォルダが存在しない場合は作成
                    if (!Directory.Exists(rdpFilesFolder))
                    {
                        Directory.CreateDirectory(rdpFilesFolder);
                    }
                    
                    // コピー先のファイル名を生成
                    var targetHost = "";
                    if (!string.IsNullOrEmpty(_currentConnection.ComputerName))
                    {
                        targetHost = _currentConnection.ComputerName;
                    }
                    else if (!string.IsNullOrEmpty(_currentConnection.IpAddressValue))
                    {
                        targetHost = _currentConnection.IpAddressValue.Replace(".", "_");
                    }
                    else if (!string.IsNullOrEmpty(_currentConnection.FullAddress))
                    {
                        var parts = _currentConnection.FullAddress.Split(':');
                        targetHost = parts[0].Replace(".", "_");
                    }
                    else
                    {
                        targetHost = System.IO.Path.GetFileNameWithoutExtension(filePath);
                    }
                    
                    // 無効な文字を除去
                    var invalidChars = System.IO.Path.GetInvalidFileNameChars();
                    foreach (var c in invalidChars)
                    {
                        targetHost = targetHost.Replace(c, '_');
                    }
                    
                    // ファイル名を生成（タイムスタンプなし）
                    var newFileName = $"{targetHost}.rdp";
                    var newFilePath = System.IO.Path.Combine(rdpFilesFolder, newFileName);
                    
                    // ファイルをコピー
                    File.Copy(filePath, newFilePath, true);
                    
                    // 現在の接続情報にrdp_filesフォルダのパスを設定
                    _currentConnection.RdpFilePath = newFilePath;
                    _currentConnection.Name = newFileName;
                    
                    LogDebug($"外部RDPファイルをコピー: {filePath} → {newFilePath}");
                }
                else
                {
                    // 既にrdp_filesフォルダ内のファイルの場合はそのまま使用
                    _currentConnection.RdpFilePath = filePath;
                    _currentConnection.Name = System.IO.Path.GetFileName(filePath);
                }
                
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
                    
                    // セッション確認を先に開始
                    var sessionCheckTask = CheckSessionInBackground(historicalConnection);
                    
                    // モニター設定を復元（メッセージボックスが出る可能性がある）
                    await RestoreMonitorSettingsAsync(historicalConnection);
                    
                    StatusText.Text = $"RDPファイルを読み込みました: {System.IO.Path.GetFileName(filePath)} (設定を履歴から復元)";
                }
                else
                {
                    // ネットワークからMACアドレスを取得を試みる
                    StatusText.Text = "MACアドレスを検索中...";
                    
                    string? macAddress = null;
                    
                    // コンピュータ名を優先してMACアドレスを取得（DHCP環境でのIP変更に対応）
                    var targetHost = !string.IsNullOrEmpty(ComputerNameTextBox.Text) ? ComputerNameTextBox.Text : IpAddressTextBox.Text;
                    
                    // コンピュータ名の場合は最新のIPアドレスを解決
                    if (!string.IsNullOrEmpty(ComputerNameTextBox.Text) && System.Net.IPAddress.TryParse(ComputerNameTextBox.Text, out _) == false)
                    {
                        await ResolveHostNameAsync(ComputerNameTextBox.Text);
                    }
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
                    
                    // バックグラウンドでセッション状態をチェック
                    _ = CheckSessionInBackground(_currentConnection);
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
                
                // コンピュータ名を優先（DHCP環境でのIP変更に対応）
                var targetHost = !string.IsNullOrEmpty(ComputerNameTextBox.Text) ? ComputerNameTextBox.Text : IpAddressTextBox.Text;
                
                // コンピュータ名の場合は最新のIPアドレスを解決してからWOL実行
                if (!string.IsNullOrEmpty(ComputerNameTextBox.Text) && System.Net.IPAddress.TryParse(ComputerNameTextBox.Text, out _) == false)
                {
                    await ResolveHostNameAsync(ComputerNameTextBox.Text);
                }
                
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
                
                // RDPファイルからの設定がある場合（UIから最新の値を取得）
                if (!string.IsNullOrEmpty(ComputerNameTextBox.Text) || !string.IsNullOrEmpty(IpAddressTextBox.Text))
                {
                    _currentConnection.ComputerName = ComputerNameTextBox.Text;
                    _currentConnection.IpAddressValue = IpAddressTextBox.Text;
                    _currentConnection.MacAddress = MacAddressTextBox.Text;
                    _currentConnection.Username = UsernameTextBox.Text; // ユーザー名も更新
                    
                    // ポート番号も最新の値を反映
                    if (!string.IsNullOrEmpty(PortTextBox.Text) && int.TryParse(PortTextBox.Text, out int port))
                    {
                        _currentConnection.Port = port;
                    }
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
                    
                    // 直打ち接続の場合、RDPファイルを自動生成
                    if (string.IsNullOrEmpty(_currentConnection.RdpFilePath))
                    {
                        var rdpFilesFolder = System.IO.Path.Combine(AppContext.BaseDirectory, "rdp_files");
                        
                        // rdp_filesフォルダが存在しない場合は作成
                        if (!Directory.Exists(rdpFilesFolder))
                        {
                            Directory.CreateDirectory(rdpFilesFolder);
                        }
                        
                        // ファイル名を生成
                        var targetHost = "";
                        if (!string.IsNullOrEmpty(_currentConnection.ComputerName))
                        {
                            targetHost = _currentConnection.ComputerName;
                        }
                        else if (!string.IsNullOrEmpty(_currentConnection.IpAddressValue))
                        {
                            targetHost = _currentConnection.IpAddressValue.Replace(".", "_");
                        }
                        else
                        {
                            targetHost = "DirectConnection";
                        }
                        
                        // 無効な文字を除去
                        var invalidChars = System.IO.Path.GetInvalidFileNameChars();
                        foreach (var c in invalidChars)
                        {
                            targetHost = targetHost.Replace(c, '_');
                        }
                        
                        // ポート番号が3389以外の場合は追加
                        if (_currentConnection.Port != 3389)
                        {
                            targetHost += $"_port{_currentConnection.Port}";
                        }
                        
                        // ファイル名を生成（タイムスタンプなし）
                        var rdpFileName = $"{targetHost}.rdp";
                        var rdpFilePath = System.IO.Path.Combine(rdpFilesFolder, rdpFileName);
                        
                        _currentConnection.RdpFilePath = rdpFilePath;
                        _currentConnection.Name = rdpFileName;
                        
                        // RDPファイルを保存
                        try
                        {
                            // モニター設定を反映
                            var selectedMonitorsForRdp = GetSelectedMonitors();
                            if (selectedMonitorsForRdp.Count > 0)
                            {
                                _currentConnection.UseMultimon = selectedMonitorsForRdp.Count > 1;
                                _currentConnection.SelectedMonitors = _monitorService.BuildSelectedMonitorsFlag(selectedMonitorsForRdp);
                                
                                // モニター設定を保存
                                _currentConnection.SavedMonitorCount = _currentMonitors.Count;
                                _currentConnection.SelectedMonitorIndices = _monitorConfigService.GetSelectedMonitorIndices(_currentMonitors);
                                _currentConnection.MonitorConfigHash = _monitorConfigService.GenerateMonitorConfigHash(_currentMonitors);
                            }
                            
                            // 画面設定を適用
                            ApplyScreenSettingsToConnection(_currentConnection);
                            
                            // FullAddressを更新
                            _currentConnection.UpdateFullAddress();
                            
                            // RDPファイルを保存
                            _rdpFileService.SaveRdpFile(rdpFilePath, _currentConnection);
                            LogDebug($"直打ち接続用RDPファイルを生成: {rdpFilePath}");
                        }
                        catch (Exception ex)
                        {
                            LogError($"RDPファイル生成エラー: {ex.Message}", ex);
                            // エラーがあっても接続は続行
                        }
                    }
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
                
                // 事前にチェック済みの結果があり、警告が必要な場合のみダイアログ表示
                if (_lastSessionCheckResult != null && _lastSessionCheckResult.IsInUseByOthers)
                {
                    // 警告ダイアログを表示
                    var dialog = new SessionWarningDialog(_lastSessionCheckResult)
                    {
                        Owner = this
                    };
                    
                    var result = dialog.ShowDialog();
                    if (!dialog.ShouldConnect)
                    {
                        StatusText.Text = "接続をキャンセルしました。";
                        return;
                    }
                }
                // チェック結果がない場合のみ実行時チェック
                else if (_lastSessionCheckResult == null)
                {
                    bool shouldConnect = await CheckSessionBeforeConnect(_currentConnection);
                    if (!shouldConnect)
                    {
                        StatusText.Text = "接続をキャンセルしました。";
                        return;
                    }
                }

                // タイマーを停止
                StopSessionCheckTimer();
                
                // 画面設定を適用
                ApplyScreenSettingsToConnection(_currentConnection);
                
                await _remoteDesktopService.ConnectAsync(_currentConnection);
                StatusText.Text = "リモートデスクトップ接続を開始しました。";
                
                // 履歴保存前にUIから最新の値を取得
                _currentConnection.Username = UsernameTextBox.Text;
                _currentConnection.MacAddress = MacAddressTextBox.Text;
                
                // ポート番号も最新の値を反映
                if (!string.IsNullOrEmpty(PortTextBox.Text) && int.TryParse(PortTextBox.Text, out int currentPort))
                {
                    _currentConnection.Port = currentPort;
                }
                
                // FullAddressを再更新
                _currentConnection.UpdateFullAddress();
                
                // 接続履歴に追加（モニター設定を含む）
                LogDebug($"Saving connection with monitor config: Count={_currentConnection.SavedMonitorCount}, Hash={_currentConnection.MonitorConfigHash}");
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
                try
                {
                    // セッション状態をリセット
                    if (SessionStatusText != null)
                    {
                        SessionStatusText.Text = "未確認";
                        SessionStatusText.Foreground = new SolidColorBrush(Colors.Gray);
                    }
                    
                    // デバッグログ出力
                    var logPath = System.IO.Path.Combine(AppContext.BaseDirectory, "debug.log");
                    var logMessage = new System.Text.StringBuilder();
                    logMessage.AppendLine($"\n=== 履歴選択デバッグ開始 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
                    logMessage.AppendLine($"selectedItem.Name: {selectedItem.Name}");
                    logMessage.AppendLine($"selectedItem.ComputerName: {selectedItem.ComputerName}");
                    logMessage.AppendLine($"selectedItem.IpAddressValue: {selectedItem.IpAddressValue}");
                    logMessage.AppendLine($"selectedItem.Port: {selectedItem.Port}");
                    logMessage.AppendLine($"selectedItem.Username: {selectedItem.Username}");
                    logMessage.AppendLine($"selectedItem.MacAddress: {selectedItem.MacAddress}");
                    logMessage.AppendLine($"selectedItem.RdpFilePath: {selectedItem.RdpFilePath}");
                    logMessage.AppendLine($"selectedItem.FullAddress: {selectedItem.FullAddress}");
                    
                    File.AppendAllText(logPath, logMessage.ToString());
                    
                    // RDPファイルが存在する場合は読み込み
                    if (!string.IsNullOrEmpty(selectedItem.RdpFilePath) && File.Exists(selectedItem.RdpFilePath))
                    {
                        logMessage.Clear();
                        logMessage.AppendLine($"RDPファイル読み込み: {selectedItem.RdpFilePath}");
                        File.AppendAllText(logPath, logMessage.ToString());
                        
                        var rdpFileService = new RdpFileService();
                        _currentConnection = rdpFileService.LoadRdpFile(selectedItem.RdpFilePath);
                        
                        // 履歴の情報で補完（RDPファイルにない情報や履歴で管理している情報）
                        _currentConnection.MacAddress = selectedItem.MacAddress;
                        _currentConnection.LastConnection = selectedItem.LastConnection;
                        _currentConnection.RdpFilePath = selectedItem.RdpFilePath;
                        
                        // ComputerNameとPortは履歴の値を優先（RDPファイルに含まれない可能性があるため）
                        if (!string.IsNullOrEmpty(selectedItem.ComputerName))
                        {
                            _currentConnection.ComputerName = selectedItem.ComputerName;
                        }
                        if (!string.IsNullOrEmpty(selectedItem.IpAddressValue))
                        {
                            _currentConnection.IpAddressValue = selectedItem.IpAddressValue;
                        }
                        // Portも履歴の値を使用（履歴で個別管理）
                        _currentConnection.Port = selectedItem.Port;
                        
                        // RDPファイルパスを表示
                        RdpFilePathTextBox.Text = selectedItem.RdpFilePath;
                        
                        StatusText.Text = "履歴のRDPファイルから詳細設定を読み込みました。";
                    }
                    else
                    {
                        logMessage.Clear();
                        logMessage.AppendLine("RDPファイルなし - 新規作成");
                        File.AppendAllText(logPath, logMessage.ToString());
                        
                        // RDPファイルがない場合は履歴データから新規作成
                        _currentConnection = selectedItem.Clone();
                        
                        // rdp_filesフォルダに新しいRDPファイルを作成
                        var rdpFilesFolder = System.IO.Path.Combine(AppContext.BaseDirectory, "rdp_files");
                        if (!Directory.Exists(rdpFilesFolder))
                        {
                            Directory.CreateDirectory(rdpFilesFolder);
                        }
                        
                        // ファイル名を生成
                        var targetHost = "";
                        if (!string.IsNullOrEmpty(_currentConnection.ComputerName))
                        {
                            targetHost = _currentConnection.ComputerName;
                        }
                        else if (!string.IsNullOrEmpty(_currentConnection.IpAddressValue))
                        {
                            targetHost = _currentConnection.IpAddressValue.Replace(".", "_");
                        }
                        else
                        {
                            targetHost = "HistoryConnection";
                        }
                        
                        // 無効な文字を除去
                        var invalidChars = System.IO.Path.GetInvalidFileNameChars();
                        foreach (var c in invalidChars)
                        {
                            targetHost = targetHost.Replace(c, '_');
                        }
                        
                        // ポート番号が3389以外の場合は追加
                        if (_currentConnection.Port != 3389)
                        {
                            targetHost += $"_port{_currentConnection.Port}";
                        }
                        
                        // ファイル名を生成（タイムスタンプなし）
                        var rdpFileName = $"{targetHost}.rdp";
                        var rdpFilePath = System.IO.Path.Combine(rdpFilesFolder, rdpFileName);
                        
                        _currentConnection.RdpFilePath = rdpFilePath;
                        _currentConnection.Name = rdpFileName;
                        
                        // RDPファイルを保存
                        try
                        {
                            var rdpFileService = new RdpFileService();
                            rdpFileService.SaveRdpFile(rdpFilePath, _currentConnection);
                            
                            // 履歴を更新してRDPファイルパスを保存
                            _historyService.UpdateConnection(_currentConnection);
                            
                            RdpFilePathTextBox.Text = rdpFilePath;
                            StatusText.Text = "履歴から設定を読み込み、新しいRDPファイルを作成しました。";
                            LogDebug($"履歴用RDPファイルを生成: {rdpFilePath}");
                        }
                        catch (Exception ex)
                        {
                            LogError($"RDPファイル生成エラー: {ex.Message}", ex);
                            RdpFilePathTextBox.Text = _currentConnection.Name ?? "";
                            StatusText.Text = "履歴から基本設定を読み込みました（RDPファイル作成失敗）。";
                        }
                    }
                    
                    // デバッグログ: _currentConnection の値
                    logMessage.Clear();
                    logMessage.AppendLine("=== _currentConnection の値 ===");
                    logMessage.AppendLine($"_currentConnection.ComputerName: {_currentConnection.ComputerName}");
                    logMessage.AppendLine($"_currentConnection.IpAddressValue: {_currentConnection.IpAddressValue}");
                    logMessage.AppendLine($"_currentConnection.Port: {_currentConnection.Port}");
                    logMessage.AppendLine($"_currentConnection.Username: {_currentConnection.Username}");
                    File.AppendAllText(logPath, logMessage.ToString());
                    
                    // UIに反映
                    ComputerNameTextBox.Text = _currentConnection.ComputerName ?? "";
                    IpAddressTextBox.Text = _currentConnection.IpAddressValue ?? "";
                    MacAddressTextBox.Text = _currentConnection.MacAddress ?? "";
                    
                    // 直接接続設定にも反映
                    logMessage.Clear();
                    logMessage.AppendLine("=== UI設定ログ ===");
                    
                    if (!string.IsNullOrEmpty(_currentConnection.ComputerName))
                    {
                        DirectAddressTextBox.Text = _currentConnection.ComputerName;
                        logMessage.AppendLine($"DirectAddressTextBox設定（ComputerName）: {_currentConnection.ComputerName}");
                    }
                    else if (!string.IsNullOrEmpty(_currentConnection.IpAddressValue))
                    {
                        DirectAddressTextBox.Text = _currentConnection.IpAddressValue;
                        logMessage.AppendLine($"DirectAddressTextBox設定（IpAddressValue）: {_currentConnection.IpAddressValue}");
                    }
                    else
                    {
                        // コンピュータ名もIPアドレスもない場合はクリア
                        DirectAddressTextBox.Text = "";
                        logMessage.AppendLine("DirectAddressTextBox設定: 空文字");
                    }
                    
                    // ポート番号は常に表示（デフォルトでも）
                    PortTextBox.Text = _currentConnection.Port.ToString();
                    logMessage.AppendLine($"PortTextBox設定: {_currentConnection.Port}");
                    
                    UsernameTextBox.Text = _currentConnection.Username ?? "";
                    logMessage.AppendLine($"UsernameTextBox設定: {_currentConnection.Username}");
                    
                    File.AppendAllText(logPath, logMessage.ToString());
                    
                    // エクスペリエンス設定をUIに反映
                    UpdateExperienceTabFromConnection(_currentConnection);
                    
                    // ローカルリソース設定をUIに反映
                    UpdateLocalResourcesTabFromConnection(_currentConnection);
                    
                    // 画面設定をUIに反映
                    UpdateScreenSettingsFromConnection(_currentConnection);
                    
                    // バックグラウンドでセッション状態をチェック（モニター設定復元より先に開始）
                    var sessionCheckTask = CheckSessionInBackground(_currentConnection);
                    
                    // モニター設定を復元（メッセージボックスが出る可能性がある）
                    await RestoreMonitorSettingsAsync(_currentConnection);
                    
                    // セッションチェックのタスクは非同期で継続
                    
                    // 非同期でPCの状態を確認（WOLインジケータ更新）
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(100); // UIの更新を待つ
                        string? targetHost = !string.IsNullOrEmpty(_currentConnection?.IpAddressValue) 
                            ? _currentConnection.IpAddressValue 
                            : _currentConnection?.ComputerName;
                        
                        if (!string.IsNullOrEmpty(targetHost))
                        {
                            await Dispatcher.InvokeAsync(async () =>
                            {
                                await CheckHostStatusAsync(targetHost);
                            });
                        }
                    });
                    
                    logMessage.Clear();
                    logMessage.AppendLine($"=== 履歴選択デバッグ終了 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===\n");
                    File.AppendAllText(logPath, logMessage.ToString());
                }
                catch (Exception ex)
                {
                    var logPath = System.IO.Path.Combine(AppContext.BaseDirectory, "debug.log");
                    var errorLog = $"\nエラー発生: {ex.Message}\nスタックトレース: {ex.StackTrace}\n";
                    File.AppendAllText(logPath, errorLog);
                    StatusText.Text = $"履歴読み込みエラー: {ex.Message}";
                }
            }
        }

        private void UpdateExperienceTabFromConnection(RdpConnection connection)
        {
            try
            {
                // パフォーマンス設定（接続タイプ）
                PerfAutoDetect.IsChecked = connection.ConnectionType == 0;
                PerfLAN.IsChecked = connection.ConnectionType == 1;
                PerfBroadband.IsChecked = connection.ConnectionType == 2;
                PerfModem.IsChecked = connection.ConnectionType == 3;
                PerfCustom.IsChecked = connection.ConnectionType == 4;
                
                // 表示設定
                DesktopBackground.IsChecked = connection.DesktopBackground;
                FontSmoothing.IsChecked = connection.FontSmoothing;
                DesktopComposition.IsChecked = connection.DesktopComposition;
                ShowWindowContents.IsChecked = connection.ShowWindowContents;
                MenuAnimations.IsChecked = connection.MenuAnimations;
                VisualStyles.IsChecked = connection.VisualStyles;
                BitmapCaching.IsChecked = connection.BitmapCaching;
                
                // その他の設定
                AutoReconnect.IsChecked = connection.AutoReconnect;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"エクスペリエンス設定のUI反映エラー: {ex.Message}");
            }
        }

        private void UpdateLocalResourcesTabFromConnection(RdpConnection connection)
        {
            try
            {
                // オーディオ設定
                // AudioMode: 0=ローカル, 1=リモート, 2=再生しない
                // UIでは適切なRadioButtonがあると仮定して実装
                
                // オーディオ録音設定
                AudioRecord.IsChecked = connection.AudioRecord;
                
                // キーボード設定
                KeyboardLocal.IsChecked = connection.KeyboardMode == 0;
                KeyboardRemote.IsChecked = connection.KeyboardMode == 1;
                KeyboardFullscreen.IsChecked = connection.KeyboardMode == 2;
                
                // ローカルデバイスとリソース
                RedirectPrinters.IsChecked = connection.RedirectPrinters;
                RedirectClipboard.IsChecked = connection.RedirectClipboard;
                RedirectSmartCards.IsChecked = connection.RedirectSmartCards;
                RedirectPorts.IsChecked = connection.RedirectPorts;
                RedirectDrives.IsChecked = connection.RedirectDrives;
                RedirectPnpDevices.IsChecked = connection.RedirectPnpDevices;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ローカルリソース設定のUI反映エラー: {ex.Message}");
            }
        }

        private async void CheckStatusButton_Click(object sender, RoutedEventArgs e)
        {
            // コンピュータ名を優先（DHCP環境でのIP変更に対応）
                var targetHost = !string.IsNullOrEmpty(ComputerNameTextBox.Text) ? ComputerNameTextBox.Text : IpAddressTextBox.Text;
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

                    // モニター構成変更の通知のみ（復元オプションは提供しない）
                    MessageBox.Show(
                        $"前回の接続時からモニター構成が変更されています。\n\n" +
                        $"{changeDescription}\n\n" +
                        $"現在のモニター構成に合わせて設定を更新します。",
                        "モニター構成の変更",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );

                    // 現在のモニター構成に合わせて設定を更新
                    if (_currentConnection != null && connection == _currentConnection)
                    {
                        _currentConnection.SavedMonitorCount = _currentMonitors?.Count ?? 0;
                        _currentConnection.SelectedMonitorIndices = _monitorConfigService.GetSelectedMonitorIndices(_currentMonitors ?? new List<MonitorInfo>());
                        _currentConnection.MonitorConfigHash = _monitorConfigService.GenerateMonitorConfigHash(_currentMonitors ?? new List<MonitorInfo>());
                        LogDebug($"Updated monitor config in connection: Count={_currentConnection.SavedMonitorCount}, Hash={_currentConnection.MonitorConfigHash}");
                    }
                }
            }

            // 利用可能なモニターの範囲内で選択を調整
            if (_currentMonitors != null && _currentMonitors.Count > 0)
            {
                if (connection.SelectedMonitorIndices != null && connection.SelectedMonitorIndices.Count > 0)
                {
                    // 現在存在するモニターのインデックスのみを保持
                    var validIndices = connection.SelectedMonitorIndices
                        .Where(idx => idx < _currentMonitors!.Count)
                        .ToList();
                    
                    if (validIndices.Count > 0)
                    {
                        // 有効なインデックスがある場合はそれを使用
                        _monitorConfigService.RestoreMonitorSelection(_currentMonitors!, validIndices);
                    }
                    else
                    {
                        // 有効なインデックスがない場合はプライマリモニターを選択
                        var primaryIndex = _currentMonitors!.FindIndex(m => m.IsPrimary);
                        if (primaryIndex >= 0)
                        {
                            _currentMonitors[primaryIndex].IsSelected = true;
                        }
                        else
                        {
                            _currentMonitors[0].IsSelected = true;
                        }
                    }
                }
                else
                {
                    // 保存された選択がない場合はプライマリモニターを選択
                    var primaryIndex = _currentMonitors!.FindIndex(m => m.IsPrimary);
                    if (primaryIndex >= 0)
                    {
                        _currentMonitors[primaryIndex].IsSelected = true;
                    }
                    else
                    {
                        _currentMonitors[0].IsSelected = true;
                    }
                }
            }
            
            // UIを更新
            if (MonitorCheckBoxList != null)
            {
                MonitorCheckBoxList.Items.Refresh();
            }
            UpdateMonitorLayout();
            
            StatusText.Text = "モニター設定を調整しました。";
            
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
                        
                        // PC状態を確認（モニター設定復元より先に開始）
                        var statusCheckTask = CheckHostStatusAsync(_currentConnection.IpAddress);
                        
                        // バックグラウンドでセッション状態もチェック
                        var sessionCheckTask = CheckSessionInBackground(_currentConnection);
                        
                        // モニター設定を復元（メッセージボックスが出る可能性がある）
                        await RestoreMonitorSettingsAsync(_currentConnection);

                        if (useWol && !string.IsNullOrEmpty(MacAddressTextBox.Text))
                        {
                            // Wake On LANを送信
                            StatusText.Text = "Wake On LANパケットを送信中...";
                            // コンピュータ名を優先（DHCP環境でのIP変更に対応）
                var targetHost = !string.IsNullOrEmpty(ComputerNameTextBox.Text) ? ComputerNameTextBox.Text : IpAddressTextBox.Text;
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
                        
                        // セッションチェックを実行
                        bool shouldConnect = await CheckSessionBeforeConnect(_currentConnection);
                        if (!shouldConnect)
                        {
                            StatusText.Text = "接続をキャンセルしました。";
                            return;
                        }
                        
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

        /// <summary>
        /// 解像度設定を初期化
        /// </summary>
        private void InitializeResolutionSettings()
        {
            try
            {
                // 標準的なRDP解像度リストを作成
                var allResolutions = new List<(int Width, int Height, string DisplayName)>
                {
                    (640, 480, "640 × 480"),
                    (800, 600, "800 × 600"),
                    (1024, 768, "1024 × 768"),
                    (1152, 864, "1152 × 864"),
                    (1280, 720, "1280 × 720 (HD)"),
                    (1280, 1024, "1280 × 1024"),
                    (1366, 768, "1366 × 768"),
                    (1440, 900, "1440 × 900"),
                    (1600, 900, "1600 × 900"),
                    (1600, 1200, "1600 × 1200"),
                    (1680, 1050, "1680 × 1050"),
                    (1920, 1080, "1920 × 1080 (Full HD)"),
                    (1920, 1200, "1920 × 1200"),
                    (2560, 1440, "2560 × 1440 (QHD)"),
                    (2560, 1600, "2560 × 1600"),
                    (3840, 2160, "3840 × 2160 (4K UHD)")
                };

                int maxWidth = 1920;
                int maxHeight = 1080;
                
                // プライマリモニターの解像度を取得
                if (_currentMonitors?.Any(m => m.IsPrimary) == true)
                {
                    var primaryMonitor = _currentMonitors.First(m => m.IsPrimary);
                    maxWidth = (int)primaryMonitor.Bounds.Width;
                    maxHeight = (int)primaryMonitor.Bounds.Height;
                    
                    // プライマリモニターの解像度より大きい解像度を除外
                    _availableResolutions = allResolutions
                        .Where(r => r.Width <= maxWidth && r.Height <= maxHeight)
                        .ToList();
                    
                    // プライマリモニターの解像度が既存リストにない場合は追加
                    if (!_availableResolutions.Any(r => r.Width == maxWidth && r.Height == maxHeight))
                    {
                        _availableResolutions.Add((maxWidth, maxHeight, $"{maxWidth} × {maxHeight} (現在)"));
                    }
                }
                else
                {
                    // モニター情報が取得できない場合は標準的な解像度のみ
                    _availableResolutions = allResolutions
                        .Where(r => r.Width <= 1920 && r.Height <= 1200)
                        .ToList();
                }
                
                // リストを解像度順にソート
                _availableResolutions = _availableResolutions
                    .OrderBy(r => r.Width * r.Height)
                    .ToList();

                // カスタム解像度をリストに追加
                _availableResolutions.Add((0, 0, "カスタム..."));
                
                // 全画面表示を最後に追加（スライダーの最右端）
                _availableResolutions.Add((-1, -1, "全画面表示"));

                // スライダーの最大値を設定
                ResolutionSlider.Maximum = _availableResolutions.Count - 1;
                
                // デフォルトでプライマリモニターの解像度を選択
                var defaultIndex = _availableResolutions.FindIndex(r => r.Width == maxWidth && r.Height == maxHeight);
                if (defaultIndex >= 0)
                {
                    ResolutionSlider.Value = defaultIndex;
                }
                else
                {
                    // プライマリモニターの解像度が見つからない場合は適切な解像度を選択
                    defaultIndex = _availableResolutions.FindIndex(r => r.Width == 1920 && r.Height == 1080);
                    if (defaultIndex >= 0)
                    {
                        ResolutionSlider.Value = defaultIndex;
                    }
                    else
                    {
                        ResolutionSlider.Value = _availableResolutions.Count - 3; // カスタムと全画面の前
                    }
                }
                
                UpdateResolutionDisplay();
            }
            catch (Exception ex)
            {
                LogError("InitializeResolutionSettings failed", ex);
            }
        }

        /// <summary>
        /// 解像度スライダーの値変更イベント
        /// </summary>
        private void ResolutionSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            UpdateResolutionDisplay();
        }

        /// <summary>
        /// 解像度表示を更新
        /// </summary>
        private void UpdateResolutionDisplay()
        {
            try
            {
                if (_availableResolutions == null || ResolutionDisplayText == null)
                    return;

                int index = (int)ResolutionSlider.Value;
                if (index >= 0 && index < _availableResolutions.Count)
                {
                    var resolution = _availableResolutions[index];
                    
                    if (resolution.Width == -1 && resolution.Height == -1) // 全画面表示
                    {
                        ResolutionDisplayText.Text = "全画面表示";
                        if (CustomResolutionGrid != null)
                            CustomResolutionGrid.Visibility = Visibility.Collapsed;
                        
                        // プライマリモニターの解像度を使用
                        if (_currentMonitors?.Any(m => m.IsPrimary) == true)
                        {
                            var primaryMonitor = _currentMonitors.First(m => m.IsPrimary);
                            if (DesktopWidthTextBox != null)
                                DesktopWidthTextBox.Text = ((int)primaryMonitor.Bounds.Width).ToString();
                            if (DesktopHeightTextBox != null)
                                DesktopHeightTextBox.Text = ((int)primaryMonitor.Bounds.Height).ToString();
                        }
                    }
                    else if (resolution.Width == 0 && resolution.Height == 0) // カスタム
                    {
                        ResolutionDisplayText.Text = "カスタム解像度";
                        if (CustomResolutionGrid != null)
                            CustomResolutionGrid.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        ResolutionDisplayText.Text = resolution.DisplayName;
                        if (CustomResolutionGrid != null)
                            CustomResolutionGrid.Visibility = Visibility.Collapsed;
                        
                        // テキストボックスの値も更新
                        if (DesktopWidthTextBox != null)
                            DesktopWidthTextBox.Text = resolution.Width.ToString();
                        if (DesktopHeightTextBox != null)
                            DesktopHeightTextBox.Text = resolution.Height.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("UpdateResolutionDisplay failed", ex);
            }
        }

        /// <summary>
        /// カスタム解像度テキストボックスの変更イベント
        /// </summary>
        private void CustomResolution_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (CustomResolutionGrid?.Visibility == Visibility.Visible)
                {
                    if (int.TryParse(DesktopWidthTextBox.Text, out int width) && 
                        int.TryParse(DesktopHeightTextBox.Text, out int height))
                    {
                        ResolutionDisplayText.Text = $"{width} × {height} px";
                    }
                    else
                    {
                        ResolutionDisplayText.Text = "カスタム解像度";
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("CustomResolution_TextChanged failed", ex);
            }
        }

        /// <summary>
        /// バックグラウンドでセッション状態をチェック
        /// </summary>
        private async Task CheckSessionInBackground(RdpConnection connection)
        {
            try
            {
                // 接続先のアドレスを取得
                string targetHost = GetTargetHost(connection);
                
                if (string.IsNullOrEmpty(targetHost))
                {
                    StopSessionCheckTimer();
                    return;
                }

                // タイマーを開始（5秒ごとに更新）
                StartSessionCheckTimer(connection);

                // ポート番号を取得
                int port = connection?.Port ?? 3389;
                
                // 初回チェックを実行
                await PerformSessionCheck(targetHost, port, connection);
            }
            catch (Exception ex)
            {
                LogError("Background session check error", ex);
            }
        }
        
        /// <summary>
        /// OSタイプを文字列からパース
        /// </summary>
        private OsType ParseOsType(string osTypeString)
        {
            if (Enum.TryParse<OsType>(osTypeString, out var osType))
            {
                return osType;
            }
            return OsType.Unknown;
        }

        /// <summary>
        /// セッションチェックタイマーを開始
        /// </summary>
        private void StartSessionCheckTimer(RdpConnection connection)
        {
            // 既存のタイマーを停止
            StopSessionCheckTimer();
            
            // 前回のセッション確認結果をクリア
            _lastSessionCheckResult = null;
            _lastSessionCheckTime = DateTime.MinValue;

            // 新しいタイマーを作成（5秒ごとに実行）
            _sessionCheckTimer = new System.Threading.Timer(async _ =>
            {
                if (!_isCheckingSession && connection != null)
                {
                    _isCheckingSession = true;
                    try
                    {
                        string targetHost = GetTargetHost(connection);
                        if (!string.IsNullOrEmpty(targetHost))
                        {
                            int port = connection?.Port ?? 3389;
                            await PerformSessionCheck(targetHost, port, connection);
                        }
                    }
                    finally
                    {
                        _isCheckingSession = false;
                    }
                }
            }, null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));
        }

        /// <summary>
        /// セッションチェックタイマーを停止
        /// </summary>
        private void StopSessionCheckTimer()
        {
            _sessionCheckTimer?.Dispose();
            _sessionCheckTimer = null;
        }

        /// <summary>
        /// セッションチェックを実行
        /// </summary>
        private async Task PerformSessionCheck(string targetHost, int port = 3389, RdpConnection? connection = null)
        {
            try
            {
                // 前回の確認結果がまだ有効な場合は、前回の結果を維持したまま新規確認を実行
                if (_lastSessionCheckResult != null && 
                    (DateTime.Now - _lastSessionCheckTime) < _sessionCheckValidDuration)
                {
                    // キャッシュが有効な間は前回の結果を表示し続ける
                    // 「確認中」を表示しない
                }
                else
                {
                    // キャッシュが無効な場合のみ「確認中」を表示
                    Dispatcher.Invoke(() => {
                        UpdateStatusIndicator(null, "セッション確認中...");
                        if (SessionStatusText != null)
                        {
                            SessionStatusText.Text = "セッション確認中...";
                            SessionStatusText.Foreground = new SolidColorBrush(Colors.Gray);
                        }
                    });
                }
                
                // YAMLからキャッシュされたOS情報を取得
                OsInfo? cachedOsInfo = null;
                if (connection != null && !string.IsNullOrEmpty(connection.CachedOsType))
                {
                    // キャッシュが30日以内の場合のみ使用
                    if (DateTime.Now - connection.CachedOsInfoTime < TimeSpan.FromDays(30))
                    {
                        cachedOsInfo = new OsInfo
                        {
                            Type = ParseOsType(connection.CachedOsType),
                            IsRdsInstalled = connection.CachedIsRdsInstalled,
                            MaxSessions = connection.CachedMaxSessions,
                            OsName = "Cached",
                            OsVersion = "Cached"
                        };
                        LogDebug($"Using cached OS info for {targetHost}: {connection.CachedOsType}");
                    }
                }
                
                // セッションチェックを実行
                var checkResult = await _sessionMonitorService.CheckSessionsAsync(targetHost, port, cachedOsInfo);
                _lastSessionCheckResult = checkResult;
                _lastSessionCheckTime = DateTime.Now;  // 確認時刻を更新
                
                // OS情報を接続履歴に保存
                if (checkResult.IsSuccess && connection != null)
                {
                    bool needUpdate = false;
                    
                    // OS情報が更新された場合、履歴に保存
                    if (checkResult?.OsInfo != null && checkResult.OsInfo.Type != OsType.Unknown)
                    {
                        string newOsType = checkResult.OsInfo.Type.ToString();
                        if (connection.CachedOsType != newOsType ||
                            connection.CachedIsRdsInstalled != checkResult.OsInfo.IsRdsInstalled ||
                            connection.CachedMaxSessions != checkResult.OsInfo.MaxSessions)
                        {
                            connection.CachedOsType = newOsType;
                            connection.CachedIsRdsInstalled = checkResult.OsInfo.IsRdsInstalled;
                            connection.CachedMaxSessions = checkResult.OsInfo.MaxSessions;
                            connection.CachedOsInfoTime = DateTime.Now;
                            needUpdate = true;
                            LogDebug($"Updated OS info cache for {targetHost}: {newOsType}");
                        }
                    }
                    
                    // 履歴を更新
                    if (needUpdate)
                    {
                        _historyService.UpdateConnection(connection);
                    }
                }
                
                // 結果に応じてセッション状態を更新
                Dispatcher.Invoke(() => {
                    if (SessionStatusText == null) return;
                    
                    if (checkResult == null || !checkResult.IsSuccess)
                    {
                        // エラー時
                        SessionStatusText.Text = "セッション確認失敗";
                        SessionStatusText.Foreground = new SolidColorBrush(Colors.Gray);
                    }
                    else if (checkResult.IsInUseByOthers)
                    {
                        // 他のユーザーが使用中
                        var otherUsers = checkResult.Sessions
                            .Where(s => s.IsActive && 
                                !s.UserName.Equals(checkResult.CurrentUser, StringComparison.OrdinalIgnoreCase))
                            .ToList();
                        
                        if (checkResult.OsInfo.WarningLevel == WarningLevel.Warning)
                        {
                            // 警告が必要（強制切断リスクあり）
                            var userNames = string.Join(", ", otherUsers.Select(u => u.FullUserName));
                            SessionStatusText.Text = $"⚠️ 警告: {userNames} が使用中\n接続すると強制切断されます";
                            SessionStatusText.Foreground = new SolidColorBrush(Colors.Red);
                        }
                        else
                        {
                            // 情報のみ（RDS有り）
                            var userNames = string.Join(", ", otherUsers.Select(u => u.FullUserName));
                            SessionStatusText.Text = $"ℹ️ 情報: {userNames} が接続中\n複数接続可能です";
                            SessionStatusText.Foreground = new SolidColorBrush(Colors.Orange);
                        }
                    }
                    else
                    {
                        // 誰も使用していない（Ping確認は別途実行）
                        Task.Run(async () => {
                            var pingResult = await _wakeOnLanService.PingHostAsync(targetHost);
                            Dispatcher.Invoke(() => {
                                if (SessionStatusText == null) return;
                                
                                if (pingResult)
                                {
                                    SessionStatusText.Text = "✅ 接続可能\n使用中のユーザーはいません";
                                    SessionStatusText.Foreground = new SolidColorBrush(Colors.Green);
                                }
                                else
                                {
                                    SessionStatusText.Text = "❌ オフライン\n対象マシンが応答しません";
                                    SessionStatusText.Foreground = new SolidColorBrush(Colors.DarkGray);
                                }
                            });
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                LogError("Session check error", ex);
                // エラー時でも前回の有効な結果があれば維持
                if (_lastSessionCheckResult != null && 
                    (DateTime.Now - _lastSessionCheckTime) < _sessionCheckValidDuration)
                {
                    // 前回の結果をそのまま維持（何もしない）
                }
            }
        }
        
        /// <summary>
        /// 接続先のホストアドレスを取得
        /// </summary>
        private string GetTargetHost(RdpConnection connection)
        {
            // コンピュータ名を優先（DHCP環境でのIP変更に対応）
            if (!string.IsNullOrEmpty(connection?.ComputerName))
            {
                return connection.ComputerName;
            }
            else if (!string.IsNullOrEmpty(connection?.FullAddress))
            {
                var parts = connection.FullAddress.Split(':');
                return parts[0];
            }
            else if (!string.IsNullOrEmpty(connection?.IpAddress))
            {
                return connection.IpAddress;
            }
            return "";
        }

        /// <summary>
        /// ステータスインジケーターのマウスオーバー時
        /// </summary>
        private void StatusIndicator_MouseEnter(object sender, MouseEventArgs e)
        {
            // 詳細なツールチップ表示（既に設定済み）
        }

        /// <summary>
        /// ステータスインジケーターのマウスアウト時
        /// </summary>
        private void StatusIndicator_MouseLeave(object sender, MouseEventArgs e)
        {
            // 特に処理なし
        }

        /// <summary>
        /// RDP接続前にセッション状態を確認
        /// </summary>
        private async Task<bool> CheckSessionBeforeConnect(RdpConnection connection)
        {
            try
            {
                // 接続先のアドレスを取得
                string targetHost = "";
                if (!string.IsNullOrEmpty(connection.FullAddress))
                {
                    // FullAddressからポート番号を除去
                    var parts = connection.FullAddress.Split(':');
                    targetHost = parts[0];
                }
                else if (!string.IsNullOrEmpty(connection.IpAddress))
                {
                    targetHost = connection.IpAddress;
                }
                else if (!string.IsNullOrEmpty(connection.ComputerName))
                {
                    targetHost = connection.ComputerName;
                }

                if (string.IsNullOrEmpty(targetHost))
                {
                    // 接続先が不明な場合はチェックをスキップ
                    return true;
                }

                StatusText.Text = "セッション状態を確認中...";
                
                // ポート番号を取得
                int portNumber = 3389;
                if (connection != null && connection.Port > 0)
                {
                    portNumber = connection.Port;
                }
                
                // セッションチェックを実行
                var checkResult = await _sessionMonitorService.CheckSessionsAsync(targetHost, portNumber);
                
                // エラーが発生した場合はそのまま接続を許可（警告なし）
                if (!checkResult.IsSuccess)
                {
                    LogDebug($"Session check failed: {checkResult.ErrorMessage}");
                    return true;
                }
                
                // 他のユーザーが使用していない場合はそのまま接続
                if (!checkResult.IsInUseByOthers)
                {
                    return true;
                }
                
                // 警告が必要な場合はダイアログを表示
                if (checkResult.OsInfo.WarningLevel == WarningLevel.Warning || 
                    (checkResult.OsInfo.WarningLevel == WarningLevel.Info && checkResult.IsInUseByOthers))
                {
                    var dialog = new SessionWarningDialog(checkResult)
                    {
                        Owner = this
                    };
                    
                    var result = dialog.ShowDialog();
                    return dialog.ShouldConnect;
                }
                
                // その他の場合は接続を許可
                return true;
            }
            catch (Exception ex)
            {
                LogError("Session check error", ex);
                // エラー時は警告なしで接続を許可
                return true;
            }
        }
        
        /// <summary>
        /// 接続設定から画面設定UIを更新
        /// </summary>
        private void UpdateScreenSettingsFromConnection(RdpConnection connection)
        {
            try
            {
                // 画面モードは解像度スライダーで管理（全画面表示はスライダーの最右端）
                
                // 解像度
                DesktopWidthTextBox.Text = connection.DesktopWidth.ToString();
                DesktopHeightTextBox.Text = connection.DesktopHeight.ToString();
                
                // 解像度スライダーを設定（履歴の解像度に対応するインデックスを検索）
                if (_availableResolutions != null && _availableResolutions.Count > 0)
                {
                    // 履歴の解像度に一致するスライダー位置を検索
                    int matchingIndex = -1;
                    for (int i = 0; i < _availableResolutions.Count; i++)
                    {
                        var res = _availableResolutions[i];
                        if (res.Width == connection.DesktopWidth && res.Height == connection.DesktopHeight)
                        {
                            matchingIndex = i;
                            break;
                        }
                    }
                    
                    if (matchingIndex >= 0)
                    {
                        ResolutionSlider.Value = matchingIndex;
                        LogDebug($"解像度スライダーを設定: インデックス {matchingIndex} ({connection.DesktopWidth}x{connection.DesktopHeight})");
                    }
                    else
                    {
                        // 一致する解像度が見つからない場合はカスタム解像度として扱う
                        // カスタム解像度は最後から3番目の位置（全画面とその他の前）
                        int customIndex = Math.Max(0, _availableResolutions.Count - 3);
                        ResolutionSlider.Value = customIndex;
                        LogDebug($"カスタム解像度として設定: インデックス {customIndex} ({connection.DesktopWidth}x{connection.DesktopHeight})");
                    }
                }
                
                // 色深度
                foreach (ComboBoxItem item in ColorDepthComboBox.Items)
                {
                    if (item.Tag?.ToString() == connection.ColorDepth.ToString())
                    {
                        ColorDepthComboBox.SelectedItem = item;
                        break;
                    }
                }
                
                // 接続タイプとパフォーマンス設定はエクスペリエンスタブで管理
                // エクスペリエンスタブの設定を更新
                if (connection.ConnectionType == 0)
                    PerfAutoDetect.IsChecked = true;
                else if (connection.ConnectionType == 1)
                    PerfLAN.IsChecked = true;
                else if (connection.ConnectionType == 2)
                    PerfBroadband.IsChecked = true;
                else if (connection.ConnectionType == 3)
                    PerfModem.IsChecked = true;
                else
                    PerfCustom.IsChecked = true;
                
                // 表示設定もエクスペリエンスタブで管理
                DesktopBackground.IsChecked = connection.DesktopBackground;
                FontSmoothing.IsChecked = connection.FontSmoothing;
                DesktopComposition.IsChecked = connection.DesktopComposition;
                ShowWindowContents.IsChecked = connection.ShowWindowContents;
                MenuAnimations.IsChecked = connection.MenuAnimations;
                VisualStyles.IsChecked = connection.VisualStyles;
                BitmapCaching.IsChecked = connection.BitmapCaching;
                AutoReconnect.IsChecked = connection.AutoReconnect;
            }
            catch (Exception ex)
            {
                LogError("画面設定更新エラー", ex);
            }
        }
        
        /// <summary>
        /// 現在のUI設定を接続設定に反映
        /// </summary>
        private void ApplyScreenSettingsToConnection(RdpConnection connection)
        {
            try
            {
                // 画面モードは解像度スライダーで管理
                // 全画面表示が選択されている場合
                int index = (int)ResolutionSlider.Value;
                if (index >= 0 && index < _availableResolutions.Count)
                {
                    var resolution = _availableResolutions[index];
                    if (resolution.Width == -1 && resolution.Height == -1)
                    {
                        connection.ScreenModeId = 2; // フルスクリーン
                    }
                    else
                    {
                        connection.ScreenModeId = 1; // ウィンドウモード（選択した解像度で表示）
                    }
                }
                
                // 解像度
                if (int.TryParse(DesktopWidthTextBox.Text, out int width))
                    connection.DesktopWidth = width;
                if (int.TryParse(DesktopHeightTextBox.Text, out int height))
                    connection.DesktopHeight = height;
                    
                // デバッグログ出力
                LogDebug($"画面設定適用: ScreenModeId={connection.ScreenModeId}, Width={connection.DesktopWidth}, Height={connection.DesktopHeight}");
                    
                // 色深度
                if (ColorDepthComboBox.SelectedItem is ComboBoxItem colorItem && 
                    int.TryParse(colorItem.Tag?.ToString(), out int colorDepth))
                    connection.ColorDepth = colorDepth;
                    
                // 接続タイプ（エクスペリエンスタブの設定から取得）
                if (PerfAutoDetect.IsChecked == true)
                    connection.ConnectionType = 0;
                else if (PerfLAN.IsChecked == true)
                    connection.ConnectionType = 1;
                else if (PerfBroadband.IsChecked == true)
                    connection.ConnectionType = 2;
                else if (PerfModem.IsChecked == true)
                    connection.ConnectionType = 3;
                else
                    connection.ConnectionType = 0; // カスタムの場合はデフォルト
                    
                // 表示設定（エクスペリエンスタブから取得）
                connection.DesktopBackground = DesktopBackground.IsChecked == true;
                connection.FontSmoothing = FontSmoothing.IsChecked == true;
                connection.DesktopComposition = DesktopComposition.IsChecked == true;
                connection.ShowWindowContents = ShowWindowContents.IsChecked == true;
                connection.MenuAnimations = MenuAnimations.IsChecked == true;
                connection.VisualStyles = VisualStyles.IsChecked == true;
                connection.BitmapCaching = BitmapCaching.IsChecked == true;
                connection.AutoReconnect = AutoReconnect.IsChecked == true;
            }
            catch (Exception ex)
            {
                LogError("画面設定適用エラー", ex);
            }
        }
    }
}