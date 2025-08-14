using System.Windows;

namespace RemoteWakeConnect.Models
{
    public class MonitorInfo
    {
        public int Index { get; set; }
        public string DeviceName { get; set; } = string.Empty;
        public bool IsPrimary { get; set; }
        public Rect Bounds { get; set; }
        public Rect WorkingArea { get; set; }
        public bool IsSelected { get; set; }
        
        // 矩形の座標プロパティ
        public int Left { get; set; }
        public int Top { get; set; }
        public int Right { get; set; }
        public int Bottom { get; set; }
        
        public int Width => (int)Bounds.Width;
        public int Height => (int)Bounds.Height;
        public int X => (int)Bounds.X;
        public int Y => (int)Bounds.Y;
        
        public string DisplayName => $"モニター {Index + 1}{(IsPrimary ? " (プライマリ)" : "")}";
        public string Resolution => $"{Width} × {Height}";
        public string Position => $"({X}, {Y})";
    }
}