using System;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using ScreenToolsLib;

namespace ScreenToolsDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // This demo controls a web browser that takes up the entire
            // second screen.
            //
            // To test it, you need to load the following url in a browser window:
            // https://www.earthcam.com/world/ireland/dublin/?cam=templebar

            Settings settings = null;

            // This code checks for a settings file and creates one if it's not there
            var settingsFilePath = $"{Environment.CurrentDirectory}\\settings.json";
            if (File.Exists(settingsFilePath))
            {
                // Uncomment to write the settings file
                //File.Delete(settingsFilePath);
            }

            if (!File.Exists(settingsFilePath))
            {
                // Settings file is not there. Create a Settings file with the demo data
                while (true)
                {
                    // Ask the user to pick a screen for the webcam saite
                    Screen screen = null;
                    int ScreenIndex = 0;
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

                        // these are just for documentation purposes, and are not used by ScreenToolsLib
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
                        // This can happen if the webcam goes out of full-screen mode
                        var textArea5 = new TextArea();
                        textArea5.Bounds = new Rectangle(24, 967, 123, 88);
                        textArea5.ComparePngPath = $"{Environment.CurrentDirectory}\\white.png";
                        textArea5.Action = UIAction.Click;
                        textArea5.Hover = true;     // Hover over the coordinates before clicking
                        textArea5.HoverMs = 1000;   // Hover for 1 second
                                                    // This shows the control bar at the bottom
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
                        textArea6.Hover = true;     // Hover over the coordinates before clicking
                        textArea6.HoverMs = 1000;   // Hover for 1 second
                                                    // This shows the control bar at the bottom
                        textArea6.ClickCoordinates = new Point()
                        {
                            X = screen.Bounds.X + 1560,
                            Y = screen.Bounds.Y + 760
                        };
                        site.TextAreas.Add(textArea6);

                        // Add the site
                        settings.Sites.Add(site);

                        // Save the file
                        var json = JsonConvert.SerializeObject(settings);
                        File.WriteAllText(settingsFilePath, json);
                        break;
                    }
                }
            }
            else
            {
                // Settings file found. Read it and deserialize
                var json = File.ReadAllText(settingsFilePath);
                settings = JsonConvert.DeserializeObject<Settings>(json);
            }

            // Run the automation
            ScreenTools.Automate(settings);
        }
    }
}
