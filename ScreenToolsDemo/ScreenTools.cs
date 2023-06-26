using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Drawing;
using IronOcr;
using System.Drawing.Imaging;
using System.IO;

public static class ScreenTools
{
    [DllImport("user32.dll")]
    static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

    private const int MOUSEEVENTF_LEFTDOWN = 0x02;
    private const int MOUSEEVENTF_LEFTUP = 0x04;

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    public struct POINT
    {
        public int X;
        public int Y;
    }

    /// <summary>
    /// Uses IronOcr to read text in an area of a bitmap
    /// </summary>
    /// <param name="bitmap"></param>
    /// <param name="rect"></param>
    /// <returns></returns>
    public static string GetTextInArea(Bitmap bitmap, Rectangle rect)
    {
        string text = string.Empty;

        // crop the bitmap to the rect
        Bitmap cropped = CropBitmap(bitmap, rect);

        // save it to a temporary file
        string tempFile = Path.GetTempFileName();
        cropped.Save(tempFile);

        // use OCR to read the text
        var Ocr = new IronTesseract();

        // read the text
        using (var Input = new OcrInput(tempFile))
        {
            var Result = Ocr.Read(Input);
            text = Result.Text;
        }

        return text;
    }

    /// <summary>
    /// Returns a cropped version of a bitmap
    /// </summary>
    /// <param name="bitmap"></param>
    /// <param name="rect"></param>
    /// <returns></returns>
    public static Bitmap CropBitmap(Bitmap bitmap, Rectangle rect)
    {
        return bitmap.Clone(rect, bitmap.PixelFormat);
    }

    /// <summary>
    /// Returns an area of a given size of a bitmap that has the highest contrast
    /// </summary>
    /// <param name="bitmap"></param>
    /// <param name="size"></param>
    /// <param name="highestStdDev"></param>
    /// <returns></returns>
    public static Rectangle FindHighestContrastRectangle(Bitmap bitmap, Size size, double highestStdDev)
    {
        Rectangle highestContrastRectangle = new Rectangle();

        for (int x = 0; x <= bitmap.Width - size.Width; x += 40)
        {
            for (int y = 0; y <= 900; y += 20)
            {
                double stdDev = ComputeStandardDeviation(x, y, bitmap, size);
                if (stdDev > highestStdDev)
                {
                    highestStdDev = stdDev;
                    highestContrastRectangle = new Rectangle(x, y, size.Width, size.Height);
                }
            }
        }

        return highestContrastRectangle;
    }

    /// <summary>
    /// Returns whether two bitmaps are equal
    /// </summary>
    /// <param name="bitmap1"></param>
    /// <param name="bitmap2"></param>
    /// <returns></returns>
    public static bool BitmapsAreEqual(Bitmap bitmap1, Bitmap bitmap2)
    {
        // Copy the bitmaps just to make sure we don't step on them
        using (var bmp1 = (Bitmap)bitmap1.Clone())
        {
            using (var bmp2 = (Bitmap)bitmap2.Clone())
            {
                var rect = new Rectangle(0, 0, bmp1.Width, bmp1.Height);

                var bmpData1 = bmp1.LockBits(rect, ImageLockMode.ReadOnly, bmp1.PixelFormat);
                var bmpData2 = bmp2.LockBits(rect, ImageLockMode.ReadOnly, bmp2.PixelFormat);

                try
                {
                    IntPtr ptr1 = bmpData1.Scan0;
                    IntPtr ptr2 = bmpData2.Scan0;

                    int bytes = Math.Abs(bmpData1.Stride) * bmp1.Height;
                    byte[] rgbValues1 = new byte[bytes];
                    byte[] rgbValues2 = new byte[bytes];

                    Marshal.Copy(ptr1, rgbValues1, 0, bytes);
                    Marshal.Copy(ptr2, rgbValues2, 0, bytes);

                    return rgbValues1.SequenceEqual(rgbValues2);
                }
                finally
                {
                    bmp1.UnlockBits(bmpData1);
                    bmp2.UnlockBits(bmpData2);
                }
            }
        }
    }

    /// <summary>
    /// Helper method to compute standard deviation
    /// </summary>
    /// <param name="x"></param>
    /// <param name="y"></param>
    /// <param name="bitmap"></param>
    /// <param name="size"></param>
    /// <returns></returns>
    public static double ComputeStandardDeviation(int x, int y, Bitmap bitmap, Size size)
    {
        int sum = 0, sumOfSquares = 0;
        for (int i = x; i < x + size.Width; i++)
        {
            for (int j = y; j < y + size.Height; j++)
            {
                Color pixel = bitmap.GetPixel(i, j);
                int pixelValue = (pixel.R + pixel.G + pixel.B) / 3; // Average to get grayscale

                sum += pixelValue;
                sumOfSquares += pixelValue * pixelValue;
            }
        }

        int numPixels = size.Width * size.Height;
        double mean = (double)sum / numPixels;
        double variance = (double)sumOfSquares / numPixels - mean * mean;

        return Math.Sqrt(variance); // Standard deviation
    }

    /// <summary>
    /// Clicks the mouse at a given coordinate
    /// </summary>
    /// <param name="x"></param>
    /// <param name="y"></param>
    public static void Click(int x, int y)
    {
        // get the current mouse position
        POINT currentMousePoint;
        GetCursorPos(out currentMousePoint);

        // move to the new location
        SetCursorPos(x, y);

        // click
        mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
        mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);

        // reset the mouse position
        SetCursorPos(currentMousePoint.X, currentMousePoint.Y);
    }

    /// <summary>
    /// returns a bitmap with adjusted contrast.
    /// </summary>
    /// <param name="image"></param>
    /// <param name="value">Positive for higher contrast. Negative for lower contrast</param>
    /// <returns></returns>
    public static Bitmap AdjustContrast(Bitmap image, float value)
    {
        float contrast = (100.0f + value) / 100.0f;
        contrast *= contrast;
        var contrastMatrix = new float[][] {
                new float[]{contrast, 0, 0, 0, 0},
                new float[]{0, contrast, 0, 0, 0},
                new float[]{0, 0, contrast, 0, 0},
                new float[]{0, 0, 0, 1.0f, 0},
                new float[]{0.001f, 0.001f, 0.001f, 0, 1},
            };

        var attributes = new ImageAttributes();
        attributes.SetColorMatrix(new ColorMatrix(contrastMatrix), ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

        var newImage = new Bitmap(image.Width, image.Height);
        using (var g = Graphics.FromImage(newImage))
        {
            g.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, attributes);
        }

        return newImage;
    }
}
