using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using RemoteWakeConnect.Models;

namespace RemoteWakeConnect.Dialogs
{
    /// <summary>
    /// セッション警告ダイアログ
    /// </summary>
    public partial class SessionWarningDialog : Window
    {
        private SessionCheckResult _checkResult;

        /// <summary>
        /// ユーザーが接続を選択したか
        /// </summary>
        public bool ShouldConnect { get; private set; } = false;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SessionWarningDialog(SessionCheckResult checkResult)
        {
            InitializeComponent();
            _checkResult = checkResult;
            SetupDialog();
        }

        /// <summary>
        /// ダイアログを設定
        /// </summary>
        private void SetupDialog()
        {
            if (_checkResult == null)
                return;

            // OS種別に応じてダイアログをカスタマイズ
            switch (_checkResult.OsInfo.Type)
            {
                case OsType.Workstation:
                    SetupWorkstationWarning();
                    break;

                case OsType.ServerWithoutRds:
                    SetupServerWithoutRdsWarning();
                    break;

                case OsType.ServerWithRds:
                    SetupServerWithRdsInfo();
                    break;

                default:
                    SetupDefaultWarning();
                    break;
            }

            // セッション一覧を設定
            if (_checkResult.Sessions != null && _checkResult.Sessions.Length > 0)
            {
                var activeSessions = _checkResult.Sessions
                    .Where(s => !string.IsNullOrEmpty(s.UserName))
                    .OrderBy(s => s.SessionId)
                    .ToList();

                if (activeSessions.Count > 0)
                {
                    SessionList.ItemsSource = activeSessions;
                    SessionDetailBorder.Visibility = Visibility.Collapsed; // 初期は非表示
                    ShowDetailsCheckBox.Visibility = Visibility.Visible;
                }
                else
                {
                    ShowDetailsCheckBox.Visibility = Visibility.Collapsed;
                }
            }
            else
            {
                ShowDetailsCheckBox.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// ワークステーションの警告設定
        /// </summary>
        private void SetupWorkstationWarning()
        {
            // ヘッダー設定
            HeaderBorder.Background = new SolidColorBrush(Color.FromRgb(255, 235, 235));
            IconText.Text = "⚠️";
            TitleText.Text = "警告: 他のユーザーが使用中です";
            SubtitleText.Text = _checkResult.OsInfo.OsName;

            // メッセージ設定
            var otherUsers = _checkResult.Sessions
                .Where(s => s.IsActive && 
                    !s.UserName.Equals(_checkResult.CurrentUser, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (otherUsers.Count > 0)
            {
                var userList = string.Join("\n", otherUsers.Select(u => $"• {u.FullUserName}"));
                MessageText.Text = $"以下のユーザーが現在使用中です:\n\n{userList}\n\n" +
                    "接続すると、上記ユーザーのセッションが強制的に切断されます。\n" +
                    "作業中のデータが失われる可能性があります。\n\n" +
                    "本当に接続しますか？";
            }
            else
            {
                MessageText.Text = _checkResult.WarningMessage;
            }

            // ボタン設定
            ConnectButton.Content = "はい、接続する";
            ConnectButton.Background = new SolidColorBrush(Color.FromRgb(255, 200, 200));
        }

        /// <summary>
        /// RDS無しサーバーの警告設定
        /// </summary>
        private void SetupServerWithoutRdsWarning()
        {
            // ヘッダー設定
            HeaderBorder.Background = new SolidColorBrush(Color.FromRgb(255, 245, 235));
            IconText.Text = "⚠️";
            TitleText.Text = "警告: 管理用セッション制限";
            SubtitleText.Text = _checkResult.OsInfo.OsName;

            // メッセージ設定
            var activeSessions = _checkResult.Sessions
                .Where(s => !string.IsNullOrEmpty(s.UserName))
                .ToList();

            MessageText.Text = "このサーバーはRDSがインストールされていないため、\n" +
                "管理用接続（最大2セッション）のみ利用可能です。\n\n" +
                $"現在 {activeSessions.Count} セッションが使用中です。\n\n" +
                "接続すると、既存のセッションが切断される可能性があります。\n\n" +
                "本当に接続しますか？";

            // ボタン設定
            ConnectButton.Content = "はい、接続する";
            ConnectButton.Background = new SolidColorBrush(Color.FromRgb(255, 230, 200));
        }

        /// <summary>
        /// RDS有りサーバーの情報設定
        /// </summary>
        private void SetupServerWithRdsInfo()
        {
            // ヘッダー設定
            HeaderBorder.Background = new SolidColorBrush(Color.FromRgb(235, 245, 255));
            IconText.Text = "ℹ️";
            TitleText.Text = "情報: 複数のユーザーが接続中です";
            SubtitleText.Text = _checkResult.OsInfo.OsName;

            // メッセージ設定
            var activeSessions = _checkResult.Sessions
                .Where(s => !string.IsNullOrEmpty(s.UserName))
                .ToList();

            MessageText.Text = "このサーバーは複数同時接続をサポートしています。\n\n" +
                $"現在 {activeSessions.Count} セッションが接続中です。\n\n" +
                "接続しても既存のセッションは維持されます。\n" +
                "安全に接続できます。";

            // 追加情報
            AdditionalInfoText.Text = "※ RDS (Remote Desktop Services) が有効になっています。";
            AdditionalInfoText.Visibility = Visibility.Visible;

            // ボタン設定
            ConnectButton.Content = "接続する";
            ConnectButton.Background = new SolidColorBrush(Color.FromRgb(200, 230, 255));
        }

        /// <summary>
        /// デフォルトの警告設定
        /// </summary>
        private void SetupDefaultWarning()
        {
            // ヘッダー設定
            HeaderBorder.Background = new SolidColorBrush(Color.FromRgb(245, 245, 245));
            IconText.Text = "⚠️";
            TitleText.Text = "セッション確認";
            SubtitleText.Text = "接続先の情報";

            // メッセージ設定
            if (!string.IsNullOrEmpty(_checkResult.WarningMessage))
            {
                MessageText.Text = _checkResult.WarningMessage;
            }
            else
            {
                MessageText.Text = "セッション情報を確認してください。";
            }

            // ボタン設定
            ConnectButton.Content = "接続する";
        }

        /// <summary>
        /// 詳細表示チェックボックス - チェック時
        /// </summary>
        private void ShowDetailsCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            SessionDetailBorder.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 詳細表示チェックボックス - チェック解除時
        /// </summary>
        private void ShowDetailsCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            SessionDetailBorder.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// 接続ボタンクリック
        /// </summary>
        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            ShouldConnect = true;
            DialogResult = true;
            Close();
        }

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            ShouldConnect = false;
            DialogResult = false;
            Close();
        }
    }
}