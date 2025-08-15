using System;
using System.Runtime.InteropServices;
using System.Text;

namespace RemoteWakeConnect.Native
{
    /// <summary>
    /// Windows Terminal Services API P/Invoke定義
    /// </summary>
    public static class WtsApi32
    {
        private const string DllName = "wtsapi32.dll";

        #region Enums
        
        /// <summary>
        /// セッション接続状態
        /// </summary>
        public enum WTS_CONNECTSTATE_CLASS
        {
            WTSActive,              // アクティブ（操作中）
            WTSConnected,           // 接続済み
            WTSConnectQuery,        // 接続クエリ中
            WTSShadow,              // シャドウセッション
            WTSDisconnected,        // 切断状態
            WTSIdle,                // アイドル状態
            WTSListen,              // リスニング
            WTSReset,               // リセット中
            WTSDown,                // ダウン
            WTSInit                 // 初期化中
        }

        /// <summary>
        /// WTS情報クラス
        /// </summary>
        public enum WTS_INFO_CLASS
        {
            WTSInitialProgram,
            WTSApplicationName,
            WTSWorkingDirectory,
            WTSOEMId,
            WTSSessionId,
            WTSUserName,
            WTSWinStationName,
            WTSDomainName,
            WTSConnectState,
            WTSClientBuildNumber,
            WTSClientName,
            WTSClientDirectory,
            WTSClientProductId,
            WTSClientHardwareId,
            WTSClientAddress,
            WTSClientDisplay,
            WTSClientProtocolType,
            WTSIdleTime,
            WTSLogonTime,
            WTSIncomingBytes,
            WTSOutgoingBytes,
            WTSIncomingFrames,
            WTSOutgoingFrames,
            WTSClientInfo,
            WTSSessionInfo
        }

        #endregion

        #region Structures

        /// <summary>
        /// セッション情報構造体
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct WTS_SESSION_INFO
        {
            public uint SessionId;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pWinStationName;
            public WTS_CONNECTSTATE_CLASS State;
        }

        /// <summary>
        /// セッション情報（Ex版）
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct WTS_SESSION_INFO_1
        {
            public uint ExecEnvId;
            public WTS_CONNECTSTATE_CLASS State;
            public uint SessionId;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pSessionName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pHostName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pUserName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pDomainName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pFarmName;
        }

        #endregion

        #region P/Invoke Methods

        /// <summary>
        /// WTSサーバーを開く
        /// </summary>
        [DllImport(DllName, SetLastError = true)]
        public static extern IntPtr WTSOpenServer([MarshalAs(UnmanagedType.LPTStr)] string pServerName);

        /// <summary>
        /// WTSサーバーを開く（Ex版）
        /// </summary>
        [DllImport(DllName, SetLastError = true)]
        public static extern IntPtr WTSOpenServerEx([MarshalAs(UnmanagedType.LPTStr)] string pServerName);

        /// <summary>
        /// ハンドルを閉じる
        /// </summary>
        [DllImport(DllName)]
        public static extern void WTSCloseServer(IntPtr hServer);

        /// <summary>
        /// セッション一覧を列挙
        /// </summary>
        [DllImport(DllName, SetLastError = true)]
        public static extern bool WTSEnumerateSessions(
            IntPtr hServer,
            [MarshalAs(UnmanagedType.U4)] uint Reserved,
            [MarshalAs(UnmanagedType.U4)] uint Version,
            out IntPtr ppSessionInfo,
            out uint pCount);

        /// <summary>
        /// セッション一覧を列挙（Ex版）
        /// </summary>
        [DllImport(DllName, SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool WTSEnumerateSessionsEx(
            IntPtr hServer,
            [MarshalAs(UnmanagedType.U4)] ref uint pLevel,
            [MarshalAs(UnmanagedType.U4)] uint Filter,
            out IntPtr ppSessionInfo,
            out uint pCount);

        /// <summary>
        /// セッション情報を照会
        /// </summary>
        [DllImport(DllName, SetLastError = true)]
        public static extern bool WTSQuerySessionInformation(
            IntPtr hServer,
            uint SessionId,
            WTS_INFO_CLASS WTSInfoClass,
            out IntPtr ppBuffer,
            out uint pBytesReturned);

        /// <summary>
        /// メモリを解放
        /// </summary>
        [DllImport(DllName)]
        public static extern void WTSFreeMemory(IntPtr pMemory);

        /// <summary>
        /// メモリを解放（Ex版）
        /// </summary>
        [DllImport(DllName, SetLastError = true)]
        public static extern bool WTSFreeMemoryEx(
            WTS_TYPE_CLASS WTSTypeClass,
            IntPtr pMemory,
            uint NumberOfEntries);

        #endregion

        #region Helper Enums

        /// <summary>
        /// WTSタイプクラス
        /// </summary>
        public enum WTS_TYPE_CLASS
        {
            WTSTypeProcessInfoLevel0,
            WTSTypeProcessInfoLevel1,
            WTSTypeSessionInfoLevel1
        }

        #endregion

        #region Constants

        public const uint WTS_CURRENT_SERVER_HANDLE = 0;
        public const uint WTS_CURRENT_SESSION = unchecked((uint)-1);

        #endregion
    }

    /// <summary>
    /// WMI用のP/Invoke定義
    /// </summary>
    public static class WmiHelper
    {
        /// <summary>
        /// Windowsプロダクトタイプ
        /// </summary>
        public enum ProductType : uint
        {
            Workstation = 1,        // ワークステーション（Windows 10/11 Pro等）
            DomainController = 2,   // ドメインコントローラー
            Server = 3              // サーバー
        }
    }
}