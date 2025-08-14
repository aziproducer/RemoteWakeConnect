using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Shell;
using RemoteWakeConnect.Services;

namespace RemoteWakeConnect
{
    public partial class App : Application
    {
        private static readonly string LogFile = Path.Combine(
            AppContext.BaseDirectory,
            "debug.log"
        );

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                // ログファイルの初期化
                File.WriteAllText(LogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Application starting...\n");
                LogDebug("OnStartup called");
                
                // コマンドライン引数を処理
                ProcessCommandLineArgs(e.Args);

                // 未処理例外のハンドリング
                AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
                DispatcherUnhandledException += App_DispatcherUnhandledException;

                LogDebug("Exception handlers registered");

                // 実行環境の情報をログ
                LogDebug($"OS Version: {Environment.OSVersion}");
                LogDebug($".NET Version: {Environment.Version}");
                LogDebug($"Working Directory: {Environment.CurrentDirectory}");
                LogDebug($"Executable Path: {AppContext.BaseDirectory}");

                base.OnStartup(e);
                LogDebug("Base OnStartup completed");

                // ジャンプリストの初期化
                InitializeJumpList();

                // メインウィンドウの作成
                try
                {
                    LogDebug("Creating MainWindow");
                    var mainWindow = new MainWindow();
                    LogDebug("MainWindow created successfully");
                    
                    mainWindow.Show();
                    LogDebug("MainWindow shown");
                    
                    // ジャンプリストを更新
                    UpdateJumpList();
                }
                catch (Exception ex)
                {
                    LogError("Failed to create MainWindow", ex);
                    MessageBox.Show(
                        $"ウィンドウの作成に失敗しました:\n{ex.Message}\n\n詳細はdebug.logを確認してください。",
                        "起動エラー",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                    Shutdown(1);
                }
            }
            catch (Exception ex)
            {
                LogError("OnStartup failed", ex);
                MessageBox.Show(
                    $"アプリケーションの起動に失敗しました:\n{ex.Message}\n\n詳細はdebug.logを確認してください。",
                    "起動エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                Shutdown(1);
            }
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            LogError("Dispatcher Unhandled Exception", e.Exception);
            
            MessageBox.Show(
                $"予期しないエラーが発生しました:\n{e.Exception.Message}\n\n詳細はdebug.logを確認してください。",
                "エラー",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );

            e.Handled = true;
            Shutdown(1);
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                LogError("AppDomain Unhandled Exception", ex);
                
                MessageBox.Show(
                    $"致命的なエラーが発生しました:\n{ex.Message}\n\n詳細はdebug.logを確認してください。",
                    "致命的エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
            else
            {
                LogDebug($"AppDomain Unhandled Exception (non-Exception): {e.ExceptionObject}");
            }
        }

        private static void LogDebug(string message)
        {
            try
            {
                File.AppendAllText(LogFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [DEBUG] {message}\n");
            }
            catch
            {
                // ログ書き込みエラーは無視
            }
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
                    errorMessage += $"  Inner StackTrace:\n{ex.InnerException.StackTrace}\n";
                }
                
                File.AppendAllText(LogFile, errorMessage);
            }
            catch
            {
                // ログ書き込みエラーは無視
            }
        }

        private void ProcessCommandLineArgs(string[] args)
        {
            if (args == null || args.Length == 0)
                return;

            LogDebug($"Processing command line args: {string.Join(" ", args)}");

            // コマンドライン引数の処理
            // 例: RemoteWakeConnect.exe /connect:192.168.1.100 /wol
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].StartsWith("/connect:"))
                {
                    var address = args[i].Substring(9);
                    Current.Properties["ConnectAddress"] = address;
                }
                else if (args[i] == "/wol")
                {
                    Current.Properties["UseWakeOnLan"] = true;
                }
                else if (args[i].StartsWith("/mac:"))
                {
                    var mac = args[i].Substring(5);
                    Current.Properties["MacAddress"] = mac;
                }
            }
        }

        private void InitializeJumpList()
        {
            try
            {
                var jumpList = new JumpList();
                jumpList.ShowFrequentCategory = false;
                jumpList.ShowRecentCategory = false;

                // カスタムカテゴリを追加
                JumpList.SetJumpList(Application.Current, jumpList);
                LogDebug("JumpList initialized");
            }
            catch (Exception ex)
            {
                LogError("Failed to initialize JumpList", ex);
            }
        }

        public void UpdateJumpList()
        {
            try
            {
                var jumpList = JumpList.GetJumpList(Application.Current);
                if (jumpList == null)
                {
                    jumpList = new JumpList();
                    JumpList.SetJumpList(Application.Current, jumpList);
                }

                // 既存のアイテムをクリア
                jumpList.JumpItems.Clear();

                // 履歴サービスから接続履歴を取得
                var historyService = new ConnectionHistoryService();
                var history = historyService.GetHistory();
                
                // 最新10件の履歴をジャンプリストに追加
                int count = 0;
                foreach (var connection in history.Take(10))
                {
                    if (string.IsNullOrEmpty(connection.FullAddress))
                        continue;

                    // 直接接続用のタスク
                    var connectTask = new JumpTask
                    {
                        Title = $"接続: {connection.FullAddress}",
                        Description = $"最終接続: {connection.LastConnection:yyyy/MM/dd HH:mm}",
                        ApplicationPath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? AppContext.BaseDirectory,
                        Arguments = $"/connect:{connection.FullAddress}",
                        CustomCategory = "最近の接続",
                        IconResourcePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? AppContext.BaseDirectory,
                        IconResourceIndex = 0
                    };
                    jumpList.JumpItems.Add(connectTask);

                    // Wake On LAN経由接続用のタスク（MACアドレスがある場合）
                    if (!string.IsNullOrEmpty(connection.MacAddress))
                    {
                        var wolTask = new JumpTask
                        {
                            Title = $"WOL→接続: {connection.FullAddress}",
                            Description = $"Wake On LAN後に接続 (MAC: {connection.MacAddress})",
                            ApplicationPath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? AppContext.BaseDirectory,
                            Arguments = $"/connect:{connection.FullAddress} /wol /mac:{connection.MacAddress}",
                            CustomCategory = "Wake On LAN",
                            IconResourcePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? AppContext.BaseDirectory,
                            IconResourceIndex = 0
                        };
                        jumpList.JumpItems.Add(wolTask);
                    }

                    count++;
                }

                // 区切り線を追加
                jumpList.JumpItems.Add(new JumpTask { CustomCategory = "" });

                // 基本タスクを追加
                jumpList.JumpItems.Add(new JumpTask
                {
                    Title = "新規接続",
                    Description = "新しい接続を開始",
                    ApplicationPath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? AppContext.BaseDirectory,
                    Arguments = "",
                    IconResourcePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? AppContext.BaseDirectory,
                    IconResourceIndex = 0
                });

                // ジャンプリストを適用
                jumpList.Apply();
                LogDebug($"JumpList updated with {count} items");
            }
            catch (Exception ex)
            {
                LogError("Failed to update JumpList", ex);
            }
        }
    }
}