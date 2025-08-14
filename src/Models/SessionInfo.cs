using System;
using RemoteWakeConnect.Native;

namespace RemoteWakeConnect.Models
{
    /// <summary>
    /// セッション情報
    /// </summary>
    public class SessionInfo
    {
        /// <summary>
        /// セッションID
        /// </summary>
        public uint SessionId { get; set; }

        /// <summary>
        /// ユーザー名
        /// </summary>
        public string UserName { get; set; } = string.Empty;

        /// <summary>
        /// ドメイン名
        /// </summary>
        public string DomainName { get; set; } = string.Empty;

        /// <summary>
        /// セッション名（Console、RDP-Tcp#1など）
        /// </summary>
        public string SessionName { get; set; } = string.Empty;

        /// <summary>
        /// セッション状態
        /// </summary>
        public WtsApi32.WTS_CONNECTSTATE_CLASS State { get; set; }

        /// <summary>
        /// ログオン時刻
        /// </summary>
        public DateTime? LogonTime { get; set; }

        /// <summary>
        /// アイドル時間
        /// </summary>
        public TimeSpan? IdleTime { get; set; }

        /// <summary>
        /// ロック状態
        /// </summary>
        public bool IsLocked { get; set; }

        /// <summary>
        /// 完全なユーザー名（ドメイン\ユーザー）
        /// </summary>
        public string FullUserName =>
            string.IsNullOrEmpty(DomainName) 
                ? UserName 
                : $"{DomainName}\\{UserName}";

        /// <summary>
        /// アクティブかどうか
        /// </summary>
        public bool IsActive =>
            State == WtsApi32.WTS_CONNECTSTATE_CLASS.WTSActive;

        /// <summary>
        /// コンソールセッションかどうか
        /// </summary>
        public bool IsConsoleSession =>
            SessionName?.Equals("Console", StringComparison.OrdinalIgnoreCase) ?? false;
    }

    /// <summary>
    /// OS情報
    /// </summary>
    public class OsInfo
    {
        /// <summary>
        /// OS種別
        /// </summary>
        public OsType Type { get; set; }

        /// <summary>
        /// RDSがインストールされているか
        /// </summary>
        public bool IsRdsInstalled { get; set; }

        /// <summary>
        /// 最大同時接続数
        /// </summary>
        public int MaxSessions { get; set; }

        /// <summary>
        /// 警告レベル
        /// </summary>
        public WarningLevel WarningLevel { get; set; }

        /// <summary>
        /// OS名
        /// </summary>
        public string OsName { get; set; } = string.Empty;

        /// <summary>
        /// OSバージョン
        /// </summary>
        public string OsVersion { get; set; } = string.Empty;
    }

    /// <summary>
    /// OS種別
    /// </summary>
    public enum OsType
    {
        /// <summary>
        /// Windows 10/11 Pro/Home
        /// </summary>
        Workstation,

        /// <summary>
        /// Windows Server（RDS無し）
        /// </summary>
        ServerWithoutRds,

        /// <summary>
        /// Windows Server（RDS有り）
        /// </summary>
        ServerWithRds,

        /// <summary>
        /// 不明
        /// </summary>
        Unknown
    }

    /// <summary>
    /// 警告レベル
    /// </summary>
    public enum WarningLevel
    {
        /// <summary>
        /// 情報（安全）
        /// </summary>
        Info,

        /// <summary>
        /// 警告（切断リスクあり）
        /// </summary>
        Warning,

        /// <summary>
        /// エラー
        /// </summary>
        Error
    }

    /// <summary>
    /// セッション確認結果
    /// </summary>
    public class SessionCheckResult
    {
        /// <summary>
        /// セッション一覧
        /// </summary>
        public SessionInfo[] Sessions { get; set; } = Array.Empty<SessionInfo>();

        /// <summary>
        /// OS情報
        /// </summary>
        public OsInfo OsInfo { get; set; } = new OsInfo();

        /// <summary>
        /// 他のユーザーが使用中か
        /// </summary>
        public bool IsInUseByOthers { get; set; }

        /// <summary>
        /// 現在のユーザー名
        /// </summary>
        public string CurrentUser { get; set; } = string.Empty;

        /// <summary>
        /// 警告メッセージ
        /// </summary>
        public string WarningMessage { get; set; } = string.Empty;

        /// <summary>
        /// 接続可能かどうか
        /// </summary>
        public bool CanConnect { get; set; }

        /// <summary>
        /// エラーメッセージ
        /// </summary>
        public string ErrorMessage { get; set; } = string.Empty;

        /// <summary>
        /// 成功したか
        /// </summary>
        public bool IsSuccess => string.IsNullOrEmpty(ErrorMessage);
    }
}