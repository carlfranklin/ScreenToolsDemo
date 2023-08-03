using System.Drawing;
namespace ScreenToolsLib
{
    public class TextArea
    {
        public string Text { get; set; } = string.Empty;
        public string ComparePngPath { get; set; } = string.Empty;
        public Rectangle Bounds { get; set; }
        public bool OnlyCheckIfFrozen { get; set; }
        public UIAction Action { get; set; }
        public int ContrastAdjustment { get; set; }
        public int RotationDegrees { get; set; }
        public string KeysToSend { get; set; } = string.Empty;
        public Point ClickCoordinates { get; set; }
        public bool Hover { get; set; }
        public int HoverMs { get; set; }
    }
}