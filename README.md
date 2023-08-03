# Screen Tools Demo

This is a demo of using a .NET Framework C# class of utilities to automate screen interactions.

The demo looks for "Still Watching?" on a streaming webcam website in a full-screen browser, and clicks on the screen when necessary. It also looks to see if the site has frozen and presses F5 to restart it. Also, if the webcam page goes out of full-screen mode, it clicks the "full-screen" button.

To use the ScreenTools in a .NET Framework Console App, follow these steps

Add references to the following assemblies by right-clicking the project and selecting **Add**, **Reference**, and clicking the **Assemblies** tab. These assemblies are part of the .NET Framework, and you must add references to them in order to use them.

Select the following assemblies (in addition to the defaults), then press **OK**.

```
System.Configuration
System.Drawing
System.Windows.Forms
```

Add a either a Project reference to **ScreenToolsLib** or an assembly reference to *ScreenToolsLib.dll*

## ScreenToolsLib

Let's look at the classes inside this library.

*UIAction.cs*:

```c#
namespace ScreenToolsLib
{
    public enum UIAction
    {
        None,
        Click,
        SendKeys
    }
}
```

The `UIAction` enumeration defines an action that you can perform. Currently you can click on a given screen coordinate and you can also send keys, simulating keyboard input.

*TextArea.cs*:

```c#
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
```

`TextArea` is the real meat of the library. It defines **what** to react to **where** on the screen, and **how** to react.

Let's look at the properties:

| **Property**       | Description                                                  |
| ------------------ | ------------------------------------------------------------ |
| Text               | Defines the text on the screen to look for                   |
| ComparePngPath     | If comparing `Bounds` to a known .PNG file, defines the path to the PNG file |
| Bounds             | Defines the rectangle (coordinates) on the screen where the text (or image) will appear |
| OnlyCheckIfFrozen  | When true, only does a check if the screen hasn't changed since the last capture |
| Action             | The `UIAction` that defines the action to take: None, Click, or SendKeys |
| ContrastAdjustment | Allows you to increase (+) or decrease (-) the contrast before comparing text |
| RotationDegrees    | When not zero, specifies the number of degrees to rotate the `Bounds` before comparing text |
| KeysToSend         | The string of keys to send. See the [documentation](https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys.send?view=windowsdesktop-7.0) for formatting. |
| ClickCoordinates   | The screen coordinates where a click should take place       |
| Hover              | When `Action` is `UIAction.Click`, and `Hover` is true, the mouse will hover over the coordinates before clicking |
| HoverMs            | The number of milliseconds to hover when `Hover` is true     |

*Site.cs*:

```c#
using System.Collections.Generic;

namespace ScreenToolsLib
{
    public class Site
    {
        public string Name { get; set; }
        public string Url { get; set; }
        public string WindowTitle { get; set; }
        public string Description { get; set; }
        public List<TextArea> TextAreas { get; set; } = new List<TextArea>();
    }
}
```

`Site` defines a website to load in a Microsoft Edge browser. The browser will be set to kiosk mode (full screen) to avoid problems with coordinates and hidden fields. 

Let's look at the properties:

| Property    | Description                                                  |
| ----------- | ------------------------------------------------------------ |
| Name        | The name of the site for documentation purposes only         |
| Url         | The address of the site                                      |
| WindowTitle | A partial or complete case-insensitive title used to identify the browser process |
| Description | A description of the site for documentation purposes only    |
| TextAreas   | The list of `TextArea` objects that define the actions to take |

*Settings.cs*:

```c#
using System.Collections.Generic;

namespace ScreenToolsLib
{
    public class Settings
    {
        public int ScreenIndex { get; set; }
        public int IntervalInSeconds { get; set; } = 5;
        public List<Site> Sites { get; set; } = new List<Site>();
    }
}
```

The `Settings` class is the top level class that defines the entirety of metadata for the actions that **ScreenToolsLib** will take.

Here are the properties:

| Property          | Description                                                  |
| ----------------- | ------------------------------------------------------------ |
| ScreenIndex       | The zero-based index of the screens as defined by Windows    |
| IntervalInSeconds | The `TextArea` objects are read and acted upon in a loop. This defines the number of seconds to pause before looping through the `TextArea` objects and acting on them. |
| Sites             | The list of Sites, and therefore the related `TextArea` objects. Currently, only one site is acted upon. |

### Using ScreenToolsLib

The easiest way to learn how to use **ScreenToolsLib** is to look at the demo *Program.cs* file. Here is the order of operations:

1. Define a `Settings` object
2. Call `ScreenTools.Automate()` passing the `Settings` object

The demo looks for *Settings.json*. If it's there, it will be loaded and deserialized, otherwise the demo code will create the `Settings` object, serialize it, and save it to *Settings.json*.

*Program.cs*:

```c#
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
            // You also need to press F11 to run browser in kiosk mode.

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
```

