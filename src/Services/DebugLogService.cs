using System;
using System.IO;

namespace RemoteWakeConnect.Services
{
    /// <summary>
    /// デバッグログサービス
    /// </summary>
    public class DebugLogService
    {
        private readonly string _logFile;

        public DebugLogService()
        {
            _logFile = Path.Combine(AppContext.BaseDirectory, "debug.log");
        }

        /// <summary>
        /// ログメッセージを書き込む
        /// </summary>
        public void WriteLine(string message)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var logMessage = $"[{timestamp}] {message}\n";
                File.AppendAllText(_logFile, logMessage);
            }
            catch
            {
                // ログ書き込みエラーは無視
            }
        }

        /// <summary>
        /// エラーログを書き込む
        /// </summary>
        public void WriteError(string message, Exception ex)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var logMessage = $"[{timestamp}] ERROR: {message}\n";
                logMessage += $"  Exception: {ex.GetType().Name}\n";
                logMessage += $"  Message: {ex.Message}\n";
                logMessage += $"  StackTrace:\n{ex.StackTrace}\n";

                if (ex.InnerException != null)
                {
                    logMessage += $"  Inner Exception: {ex.InnerException.GetType().Name}\n";
                    logMessage += $"  Inner Message: {ex.InnerException.Message}\n";
                }

                File.AppendAllText(_logFile, logMessage);
            }
            catch
            {
                // ログ書き込みエラーは無視
            }
        }
    }
}