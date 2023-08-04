using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Drawing;
using Tesseract;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms;

namespace ScreenToolsLib
{
    public static class ScreenTools
    {
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;

        public struct POINT
        {
            public int X;
            public int Y;
        }

        public static void Automate(Settings settings)
        {
            // we need the process, so we can get the window handle
            Process thisProcess = LoadOrFindBrowser(settings);

            // we need this to see if the screen has frozen
            Bitmap lastScreenShot = null;

            // get a reference to the screen
            var screen = Screen.AllScreens[settings.ScreenIndex];

            // pull out the first (and only) site
            var site = settings.Sites[0];

            // let's do this!
            Console.WriteLine($"Started monitoring at {DateTime.Now.ToShortTimeString()}");

            // automate!
            while (!Console.KeyAvailable)
            {
                // Pause
                Thread.Sleep(1000 * settings.IntervalInSeconds);

                // create a bitmap the size of the entire screen
                using (Bitmap bitmap = new Bitmap(screen.Bounds.Width,
                    screen.Bounds.Height))
                {
                    // create a new graphics object from the bitmap
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        try
                        {
                            // take the screenshot (entire screen)
                            graphics.CopyFromScreen(screen.Bounds.Location, Point.Empty,
                                screen.Bounds.Size);

                            foreach (var area in site.TextAreas)
                            {
                                if (area.OnlyCheckIfFrozen)
                                {
                                    if (lastScreenShot == null)
                                    {
                                        continue;
                                    }

                                    // is this the same as the last screen shot?
                                    if (!BitmapsAreEqual(bitmap, lastScreenShot))
                                    {
                                        lastScreenShot.Dispose();
                                        continue;
                                    }
                                }

                                // crop 
                                var cropped = CropBitmap(bitmap, area.Bounds);

                                // rotate?
                                if (area.RotationDegrees != 0)
                                {
                                    Bitmap rotated = null;
                                    switch (area.RotationDegrees)
                                    {
                                        case 90:
                                            rotated = RotateBitmap(cropped, 
                                                RotateFlipType.Rotate90FlipNone);
                                            cropped.Dispose();
                                            cropped = rotated;
                                            break;
                                        case 180:
                                            rotated = RotateBitmap(cropped, 
                                                RotateFlipType.Rotate180FlipNone);
                                            cropped.Dispose();
                                            cropped = rotated;
                                            break;
                                        case 270:
                                            rotated = RotateBitmap(cropped, 
                                                RotateFlipType.Rotate270FlipNone);
                                            cropped.Dispose();
                                            cropped = rotated;
                                            break;
                                    }
                                }

                                // contrast?
                                if (area.ContrastAdjustment != 0)
                                {
                                    var adjusted = AdjustContrast(cropped, area.ContrastAdjustment);
                                    cropped.Dispose();
                                    cropped = adjusted;
                                }

                                // compare png?
                                if (area.Text == string.Empty && area.ComparePngPath 
                                    != string.Empty)
                                {
                                    // load bitmap
                                    var patch = new Bitmap(area.ComparePngPath);

                                    // compare
                                    if (BitmapsAreEqual(cropped, patch))
                                    {
                                        ProcessUIAction(thisProcess.MainWindowHandle, area);
                                    }
                                }
                                else
                                {
                                    // get text
                                    var text = GetText(cropped);
                                    if (text.ToLower() == area.Text.ToLower())
                                    {
                                        ProcessUIAction(thisProcess.MainWindowHandle, area);
                                    }
                                }
                                cropped.Dispose();
                            }
                        }
                        catch (Exception ex)
                        {
                            // Sometimes you'll get an exception
                            // that you can ignore until the next iteration.
                        }
                    }

                    // save the bitmap
                    lastScreenShot = (Bitmap)bitmap.Clone();
                }
            }
        }

        /// <summary>
        /// Processes either a Click (with the hover feature) or a SendKeys action
        /// </summary>
        /// <param name="browserWindowHandle">The handle for the Edge Browser process</param>
        /// <param name="currentWindowHandle">The handle for the window that currently has focus</param>
        /// <param name="area"></param>
        public static void ProcessUIAction(IntPtr browserWindowHandle, TextArea area)
        {
            IntPtr currentWindowHandle = GetForegroundWindow();

            // We have a match!
            if (area.Action == UIAction.SendKeys)
            {
                // change to the browser
                SetForegroundWindow(browserWindowHandle);

                // wait .3 seconds
                Thread.Sleep(300);

                // send keys
                SendKeys.SendWait(area.KeysToSend);

                // reset the foreground window
                SetForegroundWindow(currentWindowHandle);
            }
            else if (area.Action == UIAction.Click)
            {
                // get the current cursor position
                var currentPoint = new POINT();
                GetCursorPos(out currentPoint);

                // change to the cursor position
                SetCursorPos(area.ClickCoordinates.X, area.ClickCoordinates.Y);
                if (area.Hover)
                {
                    // move the mouse by 1 pixel
                    SetCursorPos(area.ClickCoordinates.X + 1, area.ClickCoordinates.Y + 1);
                    // wait 
                    Thread.Sleep(area.HoverMs);
                }

                // click 
                ScreenTools.Click(area.ClickCoordinates.X, area.ClickCoordinates.Y);

                // reset our cursor position to where we were
                SetCursorPos(currentPoint.X, currentPoint.Y);

                // reset the foreground window
                SetForegroundWindow(currentWindowHandle);
            }
        }

        /// <summary>
        /// High-level method to look for the site in Microsoft Edge.
        /// If it's not found it is loaded, moved to the given screen index,
        /// and run in kiosk mode {F11}
        /// </summary>
        /// <param name="settings">Settings file defined by the user</param>
        /// <returns></returns>
        public static Process LoadOrFindBrowser(Settings settings)
        {
            // get the current window handle
            var currentWindowHandle = GetForegroundWindow();

            var currentSite = settings.Sites[0];

            Process thisProcess = FindProcessByWindowTitle(currentSite.WindowTitle);

            // bail if we can't find the process.
            if (thisProcess == null)
            {
                // start the process
                var edge = new Process();
                edge.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                Process.Start($"microsoft-edge:{currentSite.Url}");
                Thread.Sleep(300);

                thisProcess = FindProcessByWindowTitle(currentSite.WindowTitle);

                if (thisProcess != null)
                {
                    var screen = Screen.AllScreens[settings.ScreenIndex];

                    MoveWindow(thisProcess.MainWindowHandle,
                        screen.Bounds.X,
                        screen.Bounds.Y,
                        screen.Bounds.Width,
                        screen.Bounds.Height, true);

                    SetForegroundWindow(thisProcess.MainWindowHandle);
                    Thread.Sleep(300);
                    SendKeys.SendWait("{F11}");
                    Thread.Sleep(300);
                    SetForegroundWindow(currentWindowHandle);

                    return thisProcess;
                }
            }

            return thisProcess;
        }

        /// <summary>
        /// Returns a running process (or null) given a partial window title
        /// </summary>
        /// <param name="title">Partial text for the Window title</param>
        /// <returns></returns>
        public static Process FindProcessByWindowTitle(string title)
        {
            // get the list of processes
            Process[] processlist = Process.GetProcesses();

            // look for the one with the text
            foreach (Process process in processlist)
            {
                if (!string.IsNullOrEmpty(process.MainWindowTitle))
                {
                    // case-insensitive partial search
                    if (process.MainWindowTitle.ToLower().Contains(title.ToLower().Trim()))
                    {
                        // found it!
                        return process;
                    }
                }
            }

            return null;
        }


        /// <summary>
        /// Uses Tesseract to read text in a bitmap
        /// </summary>
        /// <param name="bitmap">The Bitmap to read from</param>
        /// <returns></returns>
        public static string GetText(Bitmap bitmap)
        {
            string text = string.Empty;

            // save it to a temporary file
            string tempFile = $"{Environment.CurrentDirectory}\\{DateTime.Now.Ticks}.tmp";
            bitmap.Save(tempFile);

            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", 
                    "eng", EngineMode.Default))
                {
                    using (var img = Pix.LoadFromFile(tempFile))
                    {
                        using (var page = engine.Process(img))
                        {
                            text = page.GetText().Trim();
                        }
                    }
                }
            }
            catch (Exception e)
            {
            }

            File.Delete(tempFile);

            return text;
        }

        /// <summary>
        /// Rotates an image
        /// </summary>
        /// <param name="bitmap">The image to rotate</param>
        /// <returns></returns>
        public static Bitmap RotateBitmap(Bitmap bitmap, RotateFlipType rotation)
        {
            var bmp = (Bitmap)bitmap.Clone();
            bmp.RotateFlip(rotation);
            return bmp;
        }

        /// <summary>
        /// Returns a cropped version of a bitmap
        /// </summary>
        /// <param name="bitmap">The Bitmap to crop</param>
        /// <param name="rect">The rectangle to crop the Bitmap to</param>
        /// <returns></returns>
        public static Bitmap CropBitmap(Bitmap bitmap, Rectangle rect)
        {
            return bitmap.Clone(rect, bitmap.PixelFormat);
        }

        /// <summary>
        /// Returns an area of a given size of a bitmap that has the highest contrast
        /// </summary>
        /// <param name="bitmap">The Bitmap to examine</param>
        /// <param name="size">The size of the rectangle to examine and return</param>
        /// <param name="highestStdDev">THe highest standard deviation allowed</param>
        /// <returns></returns>
        public static Rectangle FindHighestContrastRectangle(Bitmap bitmap, Size size, 
            double highestStdDev)
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
        /// <param name="bitmap1">The first Bitmap to compare</param>
        /// <param name="bitmap2">The second Bitmap to compare</param>
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
        /// <param name="x">X coordinate</param>
        /// <param name="y">Y coordinate</param>
        /// <param name="bitmap">The Bitmap to examine</param>
        /// <param name="size">The Size inside the bitmap</param>
        /// <returns></returns>
        public static double ComputeStandardDeviation(int x, int y, Bitmap bitmap, Size size)
        {
            int sum = 0, sumOfSquares = 0;
            for (int i = x; i < x + size.Width; i++)
            {
                for (int j = y; j < y + size.Height; j++)
                {
                    Color pixel = bitmap.GetPixel(i, j);
                    // Average to get grayscale
                    int pixelValue = (pixel.R + pixel.G + pixel.B) / 3;

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
        /// <param name="x">X coordinate</param>
        /// <param name="y">Y coordinate</param>
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
        /// Returns a bitmap with adjusted contrast.
        /// </summary>
        /// <param name="image">the Bitmap to adjust</param>
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
            attributes.SetColorMatrix(new ColorMatrix(contrastMatrix), 
                ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

            var newImage = new Bitmap(image.Width, image.Height);
            using (var g = Graphics.FromImage(newImage))
            {
                g.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height), 0, 0, 
                    image.Width, image.Height, GraphicsUnit.Pixel, attributes);
            }

            return newImage;
        }
    }
}