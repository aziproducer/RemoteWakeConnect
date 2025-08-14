using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using RemoteWakeConnect.Models;

namespace RemoteWakeConnect.Services
{
    public class MonitorService
    {
        [DllImport("user32.dll")]
        private static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct MONITORINFOEX
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public int dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szDevice;
        }

        private const uint MONITORINFOF_PRIMARY = 1;

        public List<MonitorInfo> GetMonitors()
        {
            // Windows APIを使用してモニター情報を取得（mstsc /lは存在しないオプション）
            var monitors = GetMonitorsFromWinAPI();
            
            // それでも取得できない場合はデフォルト値を返す
            if (monitors.Count == 0)
            {
                monitors.Add(new MonitorInfo
                {
                    Index = 0,
                    DeviceName = "デフォルトモニター",
                    IsPrimary = true,
                    Bounds = new Rect(0, 0, 1920, 1080),
                    WorkingArea = new Rect(0, 0, 1920, 1080),
                    Left = 0,
                    Top = 0,
                    Right = 1920,
                    Bottom = 1080
                });
            }
            
            return monitors.OrderBy(m => m.X).ThenBy(m => m.Y).ToList();
        }

        private List<MonitorInfo> GetMonitorsFromMstsc()
        {
            var monitors = new List<MonitorInfo>();
            
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = @"C:\Windows\System32\mstsc.exe",
                        Arguments = "/l",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        StandardOutputEncoding = System.Text.Encoding.Default,
                        StandardErrorEncoding = System.Text.Encoding.Default
                    }
                };

                process.Start();
                
                // タイムアウト付きで出力を読み取る
                string output = string.Empty;
                string error = string.Empty;
                
                var outputTask = Task.Run(() => process.StandardOutput.ReadToEnd());
                var errorTask = Task.Run(() => process.StandardError.ReadToEnd());
                
                if (Task.WaitAll(new[] { outputTask, errorTask }, 3000))
                {
                    output = outputTask.Result;
                    error = errorTask.Result;
                    process.WaitForExit(1000);
                    
                    if (!string.IsNullOrEmpty(error))
                    {
                        System.Diagnostics.Debug.WriteLine($"mstsc /l error: {error}");
                    }
                }
                else
                {
                    // タイムアウトした場合はプロセスを強制終了
                    try { process.Kill(); } catch { }
                    System.Diagnostics.Debug.WriteLine("mstsc /l timed out");
                    return monitors;
                }
                
                if (string.IsNullOrEmpty(output))
                {
                    System.Diagnostics.Debug.WriteLine("mstsc /l returned no output");
                    // エラー出力を試す
                    if (!string.IsNullOrEmpty(error))
                    {
                        output = error;
                    }
                    else
                    {
                        return monitors;
                    }
                }

                // デバッグ用にログ出力
                System.Diagnostics.Debug.WriteLine($"mstsc /l output:\n{output}");
                System.IO.File.WriteAllText(System.IO.Path.Combine(AppContext.BaseDirectory, "mstsc_output.log"), output);

                // GPT-5の提案に基づく複数パターン対応の解析
                monitors = ParseMstscOutput(output);
            }
            catch (Exception ex)
            {
                // エラーが発生した場合は空のリストを返す
                System.Diagnostics.Debug.WriteLine($"mstsc /l failed: {ex.Message}");
            }
            
            return monitors;
        }

        private List<MonitorInfo> ParseMstscOutput(string text)
        {
            var result = new List<MonitorInfo>();
            
            // パターン1: Display/ディスプレイ/Monitor/モニター N: 形式のヘッダー
            var headerRegex = new Regex(@"^(?:Display|ディスプレイ|Monitor|モニター)\s*(\d+)\s*[:：]", 
                RegexOptions.Multiline | RegexOptions.IgnoreCase);
            
            // パターン2: 矩形情報（Left/Top/Right/Bottom or 左/上/右/下）
            var rectLTRBRegex = new Regex(
                @"(?:Left|左)\s*=\s*(-?\d+).*?(?:Top|上)\s*=\s*(-?\d+).*?(?:Right|右)\s*=\s*(-?\d+).*?(?:Bottom|下)\s*=\s*(-?\d+)", 
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            
            // パターン3: 代替矩形情報（X/Y/Width/Height）
            var rectXYWHRegex = new Regex(
                @"(?:X|左)\s*=\s*(-?\d+).*?(?:Y|上)\s*=\s*(-?\d+).*?(?:Width|幅|横)\s*=\s*(\d+).*?(?:Height|高さ|縦)\s*=\s*(\d+)", 
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            
            // パターン4: 簡易形式 (index: width x height; (left, top, right, bottom))
            var simpleFormatRegex = new Regex(
                @"(\d+):\s*(\d+)\s*x\s*(\d+);\s*\((-?\d+),\s*(-?\d+),\s*(-?\d+),\s*(-?\d+)\)([^\r\n]*)", 
                RegexOptions.Multiline);
            
            // プライマリモニター判定
            var primaryRegex = new Regex(
                @"(?:Primary(?:\s*Monitor)?|プライマリ(?:\s*モニター)?)\s*[:：]?\s*(?:Yes|はい|True|有)|\((?:Primary|プライマリ)\)", 
                RegexOptions.IgnoreCase);
            
            // デバイス名
            var deviceRegex = new Regex(@"\\\\\.\\DISPLAY\d+", RegexOptions.IgnoreCase);

            // まず簡易形式を試す（最も一般的）
            var simpleMatches = simpleFormatRegex.Matches(text);
            if (simpleMatches.Count > 0)
            {
                foreach (Match match in simpleMatches)
                {
                    int index = int.Parse(match.Groups[1].Value);
                    int width = int.Parse(match.Groups[2].Value);
                    int height = int.Parse(match.Groups[3].Value);
                    int left = int.Parse(match.Groups[4].Value);
                    int top = int.Parse(match.Groups[5].Value);
                    int right = int.Parse(match.Groups[6].Value);
                    int bottom = int.Parse(match.Groups[7].Value);
                    string extraInfo = match.Groups[8].Value;
                    
                    var monitor = new MonitorInfo
                    {
                        Index = index,
                        DeviceName = $"\\\\\\\\.\\\\DISPLAY{index + 1}",
                        IsPrimary = primaryRegex.IsMatch(extraInfo),
                        Bounds = new Rect(left, top, right - left + 1, bottom - top + 1),
                        WorkingArea = new Rect(left, top, right - left + 1, bottom - top + 1)
                    };
                    result.Add(monitor);
                }
            }
            // ヘッダー形式を試す
            else
            {
                var headers = headerRegex.Matches(text);
                if (headers.Count > 0)
                {
                    for (int i = 0; i < headers.Count; i++)
                    {
                        var headerMatch = headers[i];
                        int idx = int.Parse(headerMatch.Groups[1].Value);
                        int start = headerMatch.Index;
                        int end = (i + 1 < headers.Count) ? headers[i + 1].Index : text.Length;
                        string block = text.Substring(start, end - start);
                        
                        var monitor = ParseMonitorBlock(block, idx, rectLTRBRegex, rectXYWHRegex, primaryRegex, deviceRegex);
                        if (monitor != null)
                        {
                            result.Add(monitor);
                        }
                    }
                }
                // ヘッダーがない場合、矩形を直接探す
                else
                {
                    var rectMatches = rectLTRBRegex.Matches(text);
                    int index = 0;
                    foreach (Match rectMatch in rectMatches)
                    {
                        var monitor = new MonitorInfo
                        {
                            Index = index++,
                            Left = int.Parse(rectMatch.Groups[1].Value),
                            Top = int.Parse(rectMatch.Groups[2].Value),
                            Right = int.Parse(rectMatch.Groups[3].Value),
                            Bottom = int.Parse(rectMatch.Groups[4].Value)
                        };
                        monitor.Bounds = new Rect(monitor.Left, monitor.Top, 
                            monitor.Right - monitor.Left, monitor.Bottom - monitor.Top);
                        monitor.WorkingArea = monitor.Bounds;
                        monitor.DeviceName = $"\\\\\\\\.\\\\DISPLAY{monitor.Index + 1}";
                        
                        // プライマリチェック（前後のテキストから判定）
                        int searchStart = Math.Max(0, rectMatch.Index - 100);
                        int searchEnd = Math.Min(text.Length, rectMatch.Index + rectMatch.Length + 100);
                        string surroundingText = text.Substring(searchStart, searchEnd - searchStart);
                        monitor.IsPrimary = primaryRegex.IsMatch(surroundingText);
                        
                        result.Add(monitor);
                    }
                }
            }
            
            return result;
        }
        
        private MonitorInfo? ParseMonitorBlock(string block, int index, 
            Regex rectLTRBRegex, Regex rectXYWHRegex, Regex primaryRegex, Regex deviceRegex)
        {
            var monitor = new MonitorInfo { Index = index };
            
            // 矩形情報を探す
            var rectMatch = rectLTRBRegex.Match(block);
            if (rectMatch.Success)
            {
                monitor.Left = int.Parse(rectMatch.Groups[1].Value);
                monitor.Top = int.Parse(rectMatch.Groups[2].Value);
                monitor.Right = int.Parse(rectMatch.Groups[3].Value);
                monitor.Bottom = int.Parse(rectMatch.Groups[4].Value);
                monitor.Bounds = new Rect(monitor.Left, monitor.Top, 
                    monitor.Right - monitor.Left, monitor.Bottom - monitor.Top);
            }
            else
            {
                // 代替形式を試す
                rectMatch = rectXYWHRegex.Match(block);
                if (rectMatch.Success)
                {
                    monitor.Left = int.Parse(rectMatch.Groups[1].Value);
                    monitor.Top = int.Parse(rectMatch.Groups[2].Value);
                    int width = int.Parse(rectMatch.Groups[3].Value);
                    int height = int.Parse(rectMatch.Groups[4].Value);
                    monitor.Right = monitor.Left + width;
                    monitor.Bottom = monitor.Top + height;
                    monitor.Bounds = new Rect(monitor.Left, monitor.Top, width, height);
                }
                else
                {
                    return null; // 矩形情報が見つからない
                }
            }
            
            monitor.WorkingArea = monitor.Bounds;
            monitor.IsPrimary = primaryRegex.IsMatch(block);
            
            // デバイス名を探す
            var deviceMatch = deviceRegex.Match(block);
            monitor.DeviceName = deviceMatch.Success ? deviceMatch.Value : $"\\\\\\\\.\\\\DISPLAY{index + 1}";
            
            return monitor;
        }

        private List<MonitorInfo> GetMonitorsFromWinAPI()
        {
            var monitors = new List<MonitorInfo>();
            int index = 0;

            MonitorEnumProc callback = (IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData) =>
            {
                var info = new MONITORINFOEX { cbSize = Marshal.SizeOf<MONITORINFOEX>() };

                if (GetMonitorInfo(hMonitor, ref info))
                {
                    var monitor = new MonitorInfo
                    {
                        Index = index++,
                        DeviceName = string.IsNullOrEmpty(info.szDevice) ? $"モニター {index}" : info.szDevice,
                        IsPrimary = (info.dwFlags & MONITORINFOF_PRIMARY) != 0,
                        Left = info.rcMonitor.Left,
                        Top = info.rcMonitor.Top,
                        Right = info.rcMonitor.Right,
                        Bottom = info.rcMonitor.Bottom,
                        Bounds = new Rect(
                            info.rcMonitor.Left,
                            info.rcMonitor.Top,
                            info.rcMonitor.Right - info.rcMonitor.Left,
                            info.rcMonitor.Bottom - info.rcMonitor.Top
                        ),
                        WorkingArea = new Rect(
                            info.rcWork.Left,
                            info.rcWork.Top,
                            info.rcWork.Right - info.rcWork.Left,
                            info.rcWork.Bottom - info.rcWork.Top
                        )
                    };
                    monitors.Add(monitor);
                }
                return true;
            };

            EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, callback, IntPtr.Zero);
            GC.KeepAlive(callback); // デリゲートをGCから保護
            
            return monitors;
        }

        public Rect GetVirtualScreenBounds()
        {
            var monitors = GetMonitors();
            if (monitors.Count == 0)
                return new Rect(0, 0, 1920, 1080);

            double left = monitors.Min(m => m.Bounds.Left);
            double top = monitors.Min(m => m.Bounds.Top);
            double right = monitors.Max(m => m.Bounds.Right);
            double bottom = monitors.Max(m => m.Bounds.Bottom);

            return new Rect(left, top, right - left, bottom - top);
        }

        public int BuildSelectedMonitorsFlag(List<MonitorInfo> selectedMonitors)
        {
            int flag = 0;
            foreach (var monitor in selectedMonitors)
            {
                flag |= (1 << monitor.Index);
            }
            return flag;
        }

        public List<int> ParseSelectedMonitorsFlag(int flag)
        {
            var indices = new List<int>();
            for (int i = 0; i < 32; i++)
            {
                if ((flag & (1 << i)) != 0)
                {
                    indices.Add(i);
                }
            }
            return indices;
        }
    }
}