﻿using System;
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

            var screen = Screen.AllScreens[settings.ScreenIndex];

            var site = settings.Sites[0];

            var currentWindowHandle = GetForegroundWindow();

            // Let's do this!
            Console.WriteLine($"Started monitoring at {DateTime.Now.ToShortTimeString()}");

            // automate!
            while (!Console.KeyAvailable)
            {
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
                                        continue;
                                    }
                                }

                                // crop 
                                var cropped = CropBitmap(bitmap, area.Bounds);

                                // rotate?
                                if (area.RotationDegrees != 0)
                                {
                                    switch (area.RotationDegrees)
                                    {
                                        case 90:
                                            cropped = RotateImage(cropped, RotateFlipType.Rotate90FlipNone);
                                            break;
                                        case 180:
                                            cropped = RotateImage(cropped, RotateFlipType.Rotate180FlipNone);
                                            break;
                                        case 270:
                                            cropped = RotateImage(cropped, RotateFlipType.Rotate270FlipNone);
                                            break;
                                    }
                                }

                                // contrast?
                                if (area.ContrastAdjustment != 0)
                                {
                                    cropped = AdjustContrast(cropped, area.ContrastAdjustment);
                                }

                                // compare png?
                                if (area.Text == string.Empty && area.ComparePngPath != string.Empty)
                                {
                                    // load bitmap
                                    var patch = new Bitmap(area.ComparePngPath);

                                    // compare
                                    if (BitmapsAreEqual(cropped, patch))
                                    {
                                        ProcessMatch(thisProcess.MainWindowHandle, currentWindowHandle, area);
                                    }
                                }
                                else
                                {
                                    // get text
                                    var text = GetTextInArea(cropped);
                                    if (text.ToLower() == area.Text.ToLower())
                                    {
                                        ProcessMatch(thisProcess.MainWindowHandle, currentWindowHandle, area);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Sometimes you'll get an exception that you can ignore until the next iteration.
                        }
                    }

                    // save the bitmap
                    lastScreenShot = (Bitmap)bitmap.Clone();
                }
            }
        }

        public static void ProcessMatch(IntPtr browserWindowHandle, IntPtr currentWindowHandle, TextArea area)
        {
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
                    // wait one second
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


        public static Process FindProcessByWindowTitle(string title)
        {
            // get the list of processes
            Process[] processlist = Process.GetProcesses();

            // look for the one with the text
            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
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
        /// Uses IronOcr to read text in an area of a bitmap
        /// </summary>
        /// <param name="bitmap"></param>
        /// <param name="rect"></param>
        /// <returns></returns>
        public static string GetTextInArea(Bitmap bitmap)
        {
            string text = string.Empty;

            // save it to a temporary file
            string tempFile = $"{Environment.CurrentDirectory}\\{DateTime.Now.Ticks}.tmp";
            bitmap.Save(tempFile);

            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
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
        /// <param name="bitmap"></param>
        /// <returns></returns>
        public static Bitmap RotateImage(Bitmap bitmap, RotateFlipType rotation)
        {
            var bmp = (Bitmap)bitmap.Clone();
            bmp.RotateFlip(rotation);
            return bmp;
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
}