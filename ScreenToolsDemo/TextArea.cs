using System.Drawing;

public class TextArea
{
    public string Text { get; set; } = string.Empty;
    public Rectangle Bounds { get; set; }
    public string ComparePngPath { get; set; } = string.Empty;
    public int ContrastAdjustment { get; set; }
    public bool OnlyCheckIfFrozen { get; set; }
    public int RotationDegrees { get; set; }
    public UIAction Action { get; set; }
    public string KeysToSend { get; set; } = string.Empty;
    public bool Hover { get; set; }
    public Point ClickCoordinates { get; set; }
}