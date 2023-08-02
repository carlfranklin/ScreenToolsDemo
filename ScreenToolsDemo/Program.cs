using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

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
        static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);


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

            Settings settings = null;
            Site currentSite = null;
            Screen screen = null;
            int ScreenIndex = 0;

            // get the current window handle
            var currentWindowHandle = GetForegroundWindow();

            // This code checks for a settings file and creates one if it's not there
            var settingsFilePath = $"{Environment.CurrentDirectory}\\settings.json";
            if (File.Exists(settingsFilePath))
            {
                // Uncomment to write the settings file
                //File.Delete(settingsFilePath);
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
                currentSite = settings.Sites[0];
                ScreenIndex = settings.ScreenIndex;
                screen = Screen.AllScreens[ScreenIndex];
            }

            // we need the process, so we can get the window handle
            Process thisProcess = ScreenTools.FindProcessByWindowTitle(currentSite.WindowTitle);

            // bail if we can't find the process.
            if (thisProcess == null)
            {
                // start the process
                var edge = new Process();
                edge.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                Process.Start($"microsoft-edge:{currentSite.Url}");
                Thread.Sleep(300);

                thisProcess = ScreenTools.FindProcessByWindowTitle(currentSite.WindowTitle);

                if (thisProcess != null)
                {
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
                }
            }

            // we need this to see if the screen has frozen
            Bitmap lastScreenShot = null;


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

                                    // compare
                                    if (ScreenTools.BitmapsAreEqual(cropped, patch))
                                    {
                                        ScreenTools.ProcessMatch(thisProcess.MainWindowHandle, currentWindowHandle, area);
                                    }
                                }
                                else
                                {
                                    // get text
                                    var text = ScreenTools.GetTextInArea(cropped);
                                    if (text.ToLower() == area.Text.ToLower())
                                    {
                                        ScreenTools.ProcessMatch(thisProcess.MainWindowHandle, currentWindowHandle, area);
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
    }
}
