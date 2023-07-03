using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using System.Diagnostics.Eventing.Reader;

namespace ScreenToolsDemo
{
    internal class Program
    {
        private static readonly object JsonSerializer;

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

        static void ProcessMatch(IntPtr browserWindowHandle, IntPtr currentWindowHandle, TextArea area)
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
                    Thread.Sleep(1000);
                }

                // click 
                ScreenTools.Click(area.ClickCoordinates.X, area.ClickCoordinates.Y);

                // reset our cursor position to where we were
                SetCursorPos(currentPoint.X, currentPoint.Y);

                // reset the foreground window
                SetForegroundWindow(currentWindowHandle);
            }
        }

        static void Main(string[] args)
        {
            // This demo controls a web browser that takes up the entire
            // second screen.
            //
            // To test it, you need to load the following url in a browser window:
            // https://www.earthcam.com/world/ireland/dublin/?cam=templebar
            // You also need to press F11 to run browser in kiosk mode.

            Settings settings = null;
            Site currentSite = null;
            Screen screen = null;

            int ScreenIndex = 0;

            // This code checks for a settings file and creates one if it's not there
            var settingsFilePath = $"{Environment.CurrentDirectory}\\settings.json";
            if (File.Exists(settingsFilePath))
            {
                File.Delete(settingsFilePath);
            }

            if (!File.Exists(settingsFilePath))
            {
                // Settings file is not there.
                while (true)
                {
                    // Ask the user to pick a screen for the webcam

                    int index = 0;
                    Console.WriteLine("Please select a secreen where the webcam browser will be shown:");
                    Console.WriteLine();
                    foreach (var scr in Screen.AllScreens)
                    {
                        Console.WriteLine($"     {index} - {scr.Bounds.Width} x {scr.Bounds.Height}");
                        index++;
                    }
                    var key = Console.ReadKey(true);
                    var keyInt = Convert.ToInt32(key.KeyChar) - 48;
                    if (keyInt < 0 || keyInt >= Screen.AllScreens.Count())
                    {
                        Console.WriteLine();
                        Console.WriteLine("Invalid Key Selected");
                        Console.WriteLine();
                        // go around again
                    }
                    else
                    {
                        // screen index selected. Save it off
                        ScreenIndex = keyInt;
                        screen = Screen.AllScreens[ScreenIndex];
                        settings = new Settings();
                        settings.ScreenIndex = ScreenIndex;

                        // set values for the demo
                        var site = new Site();

                        // this can be a partial string, used to identify the process
                        site.WindowTitle = "EarthCam";
                        site.Description = "A live webcam showing Temple Bar in Dublin";
                        site.Name = "Dublin EarthCam";
                        site.Url = "https://www.earthcam.com/world/ireland/dublin/?cam=templebar";

                        // Create the rectangle for "Still Watching?"
                        var textArea1 = new TextArea();
                        textArea1.Bounds = new Rectangle(855, 615, 190, 60);
                        textArea1.Text = "Still Watching?";
                        textArea1.Action = UIAction.Click;
                        textArea1.ClickCoordinates = new Point()
                        {
                            // Click here
                            X = screen.Bounds.X + screen.Bounds.Width / 2,
                            Y = screen.Bounds.Y + screen.Bounds.Height / 2
                        };
                        site.TextAreas.Add(textArea1);

                        // Create the rectangle for "NINE", which needs rotated
                        var textArea2 = new TextArea();
                        textArea2.OnlyCheckIfFrozen = true;
                        textArea2.Bounds = new Rectangle(99, 32, 35, 64);
                        textArea2.Text = "Nine";
                        textArea2.RotationDegrees = 90; // Rotate 90 degrees right
                        textArea2.Action = UIAction.SendKeys;
                        textArea2.KeysToSend = "{F5}";
                        site.TextAreas.Add(textArea2);

                        // Create a rectangle for "NINE", which needs rotated, and adjust contrast
                        var textArea3 = new TextArea();
                        textArea3.OnlyCheckIfFrozen = true;
                        textArea3.Bounds = new Rectangle(99, 32, 35, 64);
                        textArea3.Text = "Nine";
                        textArea3.RotationDegrees = 90;
                        textArea3.ContrastAdjustment = -50; // adjust contrast 
                        textArea3.Action = UIAction.SendKeys;
                        textArea3.KeysToSend = "{F5}";
                        site.TextAreas.Add(textArea3);

                        // Create a rectangle for "Dublin Cam"
                        var textArea4 = new TextArea();
                        textArea4.OnlyCheckIfFrozen = true;
                        textArea4.Bounds = new Rectangle(20, 940, 137, 31);
                        textArea4.Text = "Dublin Cam";
                        textArea4.Action = UIAction.SendKeys;
                        textArea4.KeysToSend = "{F5}";
                        site.TextAreas.Add(textArea4);

                        // Look for a specific png in a specific area
                        var textArea5 = new TextArea();
                        textArea5.Bounds = new Rectangle(24, 967, 123, 88);
                        textArea5.ComparePngPath = $"{Environment.CurrentDirectory}\\white.png";
                        textArea5.Action = UIAction.Click;
                        textArea5.Hover = true; // Hover over the coordinates before clicking
                        textArea5.ClickCoordinates = new Point()
                        {
                            X = screen.Bounds.X + 1560,
                            Y = screen.Bounds.Y + 760
                        };
                        site.TextAreas.Add(textArea5);

                        // Look for text indicating the page has reset from full-screen mode
                        var textArea6 = new TextArea();
                        textArea6.Bounds = new Rectangle(335, 805, 440, 265);
                        textArea6.Text = "Welcome to Dublin, Ireland";
                        textArea6.Action = UIAction.Click;
                        textArea6.Hover = true; // Hover over the coordinates before clicking
                        textArea6.ClickCoordinates = new Point()
                        {
                            X = screen.Bounds.X + 1560,
                            Y = screen.Bounds.Y + 760
                        };
                        site.TextAreas.Add(textArea6);

                        // Add the sites
                        settings.Sites.Add(site);

                        // set the site
                        currentSite = site;

                        // Save the file
                        var json = JsonConvert.SerializeObject(settings);
                        File.WriteAllText(settingsFilePath, json);
                        break;
                    }
                }
            }
            else
            {
                // Settings file found. Read the ScreenIndex
                var json = File.ReadAllText(settingsFilePath);
                settings = JsonConvert.DeserializeObject<Settings>(json);
                ScreenIndex = settings.ScreenIndex;
                screen = Screen.AllScreens[ScreenIndex];
            }

            // we need the process, so we can get the window handle
            Process thisProcess = null;

            // we need this to see if the screen has frozen
            Bitmap lastScreenShot = null;

            // get the list of processes
            Process[] processlist = Process.GetProcesses();

            // get the current window handle
            var currentWindowHandle = GetForegroundWindow();

            // look for earthcam
            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    if (process.MainWindowTitle.ToLower().Contains(currentSite.WindowTitle.ToLower()))
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
                // start the process
            }

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

                            foreach (var area in currentSite.TextAreas)
                            {
                                if (area.OnlyCheckIfFrozen)
                                {
                                    if (lastScreenShot == null)
                                    {
                                        continue;
                                    }

                                    // is this the same as the last screen shot?
                                    if (!ScreenTools.BitmapsAreEqual(bitmap, lastScreenShot))
                                    {
                                        continue;
                                    }
                                }

                                // crop 
                                var cropped = ScreenTools.CropBitmap(bitmap, area.Bounds);

                                // rotate?
                                if (area.RotationDegrees == 90)
                                {
                                    cropped = ScreenTools.RotateImageRight90Degrees(cropped);
                                }

                                // contrast?
                                if (area.ContrastAdjustment != 0)
                                {
                                    cropped = ScreenTools.AdjustContrast(cropped, area.ContrastAdjustment);
                                }

                                // compare png?
                                if (area.Text == string.Empty && area.ComparePngPath != string.Empty)
                                {
                                    // load bitmap
                                    var patch = new Bitmap(area.ComparePngPath);

                                    var testBmpPath = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\test.bmp";
                                    patch.Save(testBmpPath);

                                    // compare
                                    if (ScreenTools.BitmapsAreEqual(cropped, patch))
                                    {
                                        ProcessMatch(thisProcess.MainWindowHandle, currentWindowHandle, area);
                                    }
                                }
                                else
                                {
                                    // get text
                                    var text = ScreenTools.GetTextInArea(cropped);
                                    if (text.ToLower() == area.Text.ToLower())
                                    {
                                        ProcessMatch(thisProcess.MainWindowHandle, currentWindowHandle, area);
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }
                    }

                    // save the bitmap
                    lastScreenShot = (Bitmap)bitmap.Clone();
                }
            }

            return;
            // hard  coded stuff

            // Keep looping until the user presses a key
            while (!Console.KeyAvailable)
            {
                // wait 5 seconds
                Thread.Sleep(5000);

                // check to see if the screen is present
                if (Screen.AllScreens.Length > 1)
                {
                    // specify the area to do OCR on
                    Rectangle rect = new Rectangle(855, 615, 190, 60);

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
                                        var text1 = ScreenTools.GetTextInArea(rotated1);
                                        var text1a = ScreenTools.GetTextInArea(highContrast);

                                        // just in case, look for the text "Dublin Cam" which is higher contrast at night
                                        var textRect2 = new Rectangle(20, 940, 137, 31);

                                        // crop
                                        var cropped2 = ScreenTools.CropBitmap(bitmap, textRect2);

                                        // read the text
                                        var rect2 = new Rectangle(0, 0, 137, 31);
                                        var text2 = ScreenTools.GetTextInArea(cropped2);

                                        // text found ?
                                        if (text1.ToLower() == "nine" ||
                                            text1a.ToLower() == "nine" ||
                                            text2.ToLower() == "dublin cam")
                                        {
                                            // yes. The webcam must have stalled.
                                            // get the current window handle
                                            currentWindowHandle = GetForegroundWindow();

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
                                    currentWindowHandle = GetForegroundWindow();

                                    // Calculate the point to hover and click
                                    int x = screen.Bounds.X + 1560;
                                    int y = screen.Bounds.Y + 760;

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
                                var text = ScreenTools.GetTextInArea(bitmap);

                                // coes the screen say "Still Watching?"
                                if (text.ToLower().Contains("still watching?"))
                                {
                                    // Yes! Send a mouse click to the middle of the screen
                                    Console.WriteLine($"Reset after 'Still Watching?' at {DateTime.Now.ToShortTimeString()}");

                                    // get the current window handle
                                    currentWindowHandle = GetForegroundWindow();

                                    // change to the browser
                                    SetForegroundWindow(thisProcess.MainWindowHandle);

                                    // calculate the middle point of secondary screen
                                    int x = screen.Bounds.X + screen.Bounds.Width / 2;
                                    int y = screen.Bounds.Y + screen.Bounds.Height / 2;

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
