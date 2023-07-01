using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;

namespace ScreenToolsDemo
{
    internal class Program
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        public struct POINT
        {
            public int X;
            public int Y;
        }

        static void Main(string[] args)
        {
            // This demo controls a web browser that takes up the entire
            // second screen.
            //
            // To test it, you need to load the following url in a browser window:
            // https://www.earthcam.com/world/ireland/dublin/?cam=templebar
            // You also need to press F11 to run browser in kiosk mode.
            
            // we need the process, so we can get the window handle
            Process thisProcess = null;

            // we need this to see if the screen has frozen
            Bitmap lastScreenShot = null;

            // get the list of processes
            Process[] processlist = Process.GetProcesses();

            // look for earthcam
            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    if (process.MainWindowTitle.ToLower().StartsWith("earthcam"))
                    {
                        // found it!
                        thisProcess = process;
                        break;
                    }
                }
            }

            // bail if we can't find the process.
            if (thisProcess == null)
            {
                Console.WriteLine("Can't find EarthCam");
                return;
            }

            // Let's do this!
            Console.WriteLine($"Started monitoring at {DateTime.Now.ToShortTimeString()}");

            // Keep looping until the user presses a key
            while (!Console.KeyAvailable)
            {
                // wait 5 seconds
                Thread.Sleep(5000);

                // check to see if a secondary screen is present
                if (Screen.AllScreens.Length > 1)
                {
                    // get the secondary screen
                    var secondaryScreen = Screen.AllScreens[1];

                    // specify the area to do OCR on
                    Rectangle rect = new Rectangle(855, 615, 190, 60);

                    // create a bitmap the size of the entire screen
                    using (Bitmap bitmap = new Bitmap(secondaryScreen.Bounds.Width, 
                        secondaryScreen.Bounds.Height))
                    {
                        // create a new graphics object from the bitmap
                        using (Graphics graphics = Graphics.FromImage(bitmap))
                        {
                            try
                            {
                                // take the screenshot (entire screen)
                                graphics.CopyFromScreen(secondaryScreen.Bounds.Location, Point.Empty,
                                    secondaryScreen.Bounds.Size);

                                // have we done this once before?
                                if (lastScreenShot != null)
                                {
                                    // check to see if this screen shot and the last screen shot
                                    // are equal. That means the app is stalled.
                                    if (ScreenTools.BitmapsAreEqual(bitmap, lastScreenShot))
                                    {
                                        // get vertical text area that identifies the webcam
                                        var textRect1 = new Rectangle(99, 32, 35, 64);

                                        // crop
                                        var cropped1 = ScreenTools.CropBitmap(bitmap, textRect1);

                                        // rotate right
                                        var rotated1 = ScreenTools.RotateImageRight90Degrees(cropped1);

                                        // get a high-contrast version
                                        var highContrast = ScreenTools.AdjustContrast(rotated1, -50);

                                        // save to test
                                       // var testBmpPath = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\test.bmp";
                                       // highContrast.Save(testBmpPath);

                                        // read the text
                                        var rect1 = new Rectangle(0, 0, 64, 35);
                                        var text1 = ScreenTools.GetTextInArea(rotated1, rect1);
                                        var text1a = ScreenTools.GetTextInArea(highContrast, rect1);

                                        // just in case, look for the text "Dublin Cam" which is higher contrast at night
                                        var textRect2 = new Rectangle(20, 940, 137, 31);

                                        // crop
                                        var cropped2 = ScreenTools.CropBitmap(bitmap, textRect2);

                                        // read the text
                                        var rect2 = new Rectangle(0, 0, 137, 31);
                                        var text2 = ScreenTools.GetTextInArea(cropped2, rect2);

                                        // text found ?
                                        if (text1.ToLower() == "nine" ||
                                            text1a.ToLower() == "nine" ||
                                            text2.ToLower() == "dublin cam")
                                        {
                                            // yes. The webcam must have stalled.
                                            // get the current window handle
                                            var currentWindowHandle = GetForegroundWindow();

                                            // change to the browser
                                            SetForegroundWindow(thisProcess.MainWindowHandle);

                                            // wait .3 seconds
                                            Thread.Sleep(300);

                                            // press F5 key
                                            SendKeys.SendWait("{F5}");

                                            // reset the foreground window
                                            SetForegroundWindow(currentWindowHandle);
                                        }
                                    }
                                }
                                
                                // save off this screen shot
                                lastScreenShot = (Bitmap)bitmap.Clone();

                                //// uncomment to save the screen to the desktop for a quick look.
                                //var screenFile = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\entireScreen.png";
                                //bitmap.Save(screenFile, System.Drawing.Imaging.ImageFormat.Png);

                                // ----- check to see if the webcam page is no longer full-screen

                                // specify the area where the screen might be white
                                var whiteRect = new Rectangle(24, 967, 123, 88);

                                // crop the area that matches a png file we have saved to disk
                                var whitePatch = ScreenTools.CropBitmap(bitmap, whiteRect);

                                // load the bitmap to compare
                                var whitePngFile = $"{Environment.CurrentDirectory}\\white.png";
                                var whitePng = new Bitmap(whitePngFile);

                                // ff this returns true, then the webcam is not full screen,
                                // which can only be set with a mouse click in the lower-right corner,
                                // and that requres us to hover over that area
                                if (ScreenTools.BitmapsAreEqual(whitePatch, whitePng))
                                {
                                    // the webcam page isn't full screen any more
                                    // get the current window handle
                                    var currentWindowHandle = GetForegroundWindow();

                                    // Calculate the point to hover and click
                                    int x = secondaryScreen.Bounds.X + 1560;
                                    int y = secondaryScreen.Bounds.Y + 760;

                                    // get the current cursor position
                                    var currentPoint = new POINT();
                                    GetCursorPos(out currentPoint);

                                    // change to the hover position
                                    SetCursorPos(x, y);

                                    // move the mouse by 1 pixel
                                    SetCursorPos(x + 1, y + 1);

                                    // wait one second
                                    Thread.Sleep(1000);
                                    
                                    // click on the fullscreen icon
                                    ScreenTools.Click(x, y);

                                    // reset our cursor position to where we were
                                    SetCursorPos(currentPoint.X, currentPoint.Y);

                                    // reset the foreground window
                                    SetForegroundWindow(currentWindowHandle);
                                }

                                // get the text in the "Still Watching?" area
                                var text = ScreenTools.GetTextInArea(bitmap, rect);

                                // coes the screen say "Still Watching?"
                                if (text.ToLower().Contains("still watching?"))
                                {
                                    // Yes! Send a mouse click to the middle of the screen
                                    Console.WriteLine($"Reset after 'Still Watching?' at {DateTime.Now.ToShortTimeString()}");
                                    
                                    // get the current window handle
                                    var currentWindowHandle = GetForegroundWindow();

                                    // change to the browser
                                    SetForegroundWindow(thisProcess.MainWindowHandle);

                                    // calculate the middle point of secondary screen
                                    int x = secondaryScreen.Bounds.X + secondaryScreen.Bounds.Width / 2;
                                    int y = secondaryScreen.Bounds.Y + secondaryScreen.Bounds.Height / 2;

                                    // click the middle point of secondary screen
                                    ScreenTools.Click(x, y);

                                    // reset the foreground window
                                    SetForegroundWindow(currentWindowHandle);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.ToString());
                            }
                        }
                    }
                }
            }

            // a key was pressed
            Console.WriteLine("Press ENTER to exit");
            Console.ReadLine();
        }
    }
}
