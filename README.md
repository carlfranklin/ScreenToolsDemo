# Screen Tools Demo

This is a demo of using a .NET Framework C# class of utilities to automate screen interactions.

The demo looks for "Still Watching?" on a streaming webcam website in a full-screen browser, and clicks on the screen when necessary. It also looks to see if the site has frozen and presses F5 to restart it. Also, if the webcam page goes out of full-screen mode, it clicks the "full-screen" button.

To use the ScreenTools in a .NET Framework Console App, follow these steps.

Add the NuGet Package `Newtonsoft.Json` to the project.

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

Read the comments where the `TextArea` objects are being created. The demo controls a webcam website, reacting to situations where the webcam pauses or freezes, keeping it going while the demo is running.

> :point_up: **Note**: This demo does not condone manipulating websites with actions that may be prohibited by the website's license and/or usage agreement(s). Please use this demo for demonstration purposes only.

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

This is the resulting *Settings.json* file:

```json
{
  "ScreenIndex": 1,
  "IntervalInSeconds": 5,
  "Sites": [
    {
      "Name": "Dublin EarthCam",
      "Url": "https://www.earthcam.com/world/ireland/dublin/?cam=templebar",
      "WindowTitle": "EarthCam",
      "Description": "A live webcam showing Temple Bar in Dublin",
      "TextAreas": [
        {
          "Text": "Still Watching?",
          "Bounds": "855, 615, 190, 60",
          "ComparePngPath": "",
          "ContrastAdjustment": 0,
          "OnlyCheckIfFrozen": false,
          "RotationDegrees": 0,
          "Action": 1,
          "KeysToSend": "",
          "Hover": false,
          "HoverMS": 0,
          "ClickCoordinates": "6080, 540"
        },
        {
          "Text": "Nine",
          "Bounds": "99, 32, 35, 64",
          "ComparePngPath": "",
          "ContrastAdjustment": 0,
          "OnlyCheckIfFrozen": true,
          "RotationDegrees": 90,
          "Action": 2,
          "KeysToSend": "{F5}",
          "Hover": false,
          "HoverMS": 0,
          "ClickCoordinates": "0, 0"
        },
        {
          "Text": "Nine",
          "Bounds": "99, 32, 35, 64",
          "ComparePngPath": "",
          "ContrastAdjustment": -50,
          "OnlyCheckIfFrozen": true,
          "RotationDegrees": 90,
          "Action": 2,
          "KeysToSend": "{F5}",
          "Hover": false,
          "HoverMS": 0,
          "ClickCoordinates": "0, 0"
        },
        {
          "Text": "Dublin Cam",
          "Bounds": "20, 940, 137, 31",
          "ComparePngPath": "",
          "ContrastAdjustment": 0,
          "OnlyCheckIfFrozen": true,
          "RotationDegrees": 0,
          "Action": 2,
          "KeysToSend": "{F5}",
          "Hover": false,
          "HoverMS": 0,
          "ClickCoordinates": "0, 0"
        },
        {
          "Text": "",
          "Bounds": "24, 967, 123, 88",
          "ComparePngPath": "white.png",
          "ContrastAdjustment": 0,
          "OnlyCheckIfFrozen": false,
          "RotationDegrees": 0,
          "Action": 1,
          "KeysToSend": "",
          "Hover": true,
          "HoverMS": 1000,
          "ClickCoordinates": "6680, 760"
        },
        {
          "Text": "Welcome to Dublin, Ireland",
          "Bounds": "335, 805, 440, 265",
          "ComparePngPath": "",
          "ContrastAdjustment": 0,
          "OnlyCheckIfFrozen": false,
          "RotationDegrees": 0,
          "Action": 1,
          "KeysToSend": "",
          "Hover": true,
          "HoverMS": 1000,
          "ClickCoordinates": "6680, 760"
        }
      ]
    }
  ]
}
```

## ScreenToolsLib Methods

Let's take a look at how the methods inside **ScreenToolsLib** work.

**API Imports:**

```c#
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
```

These statements, constants, and structs are required to access the Windows API to perform required actions.

**Automate**:

```c#
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

    // Let's do this!
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
                                    cropped = RotateBitmap(cropped, 
                                        RotateFlipType.Rotate90FlipNone);
                                    break;
                                case 180:
                                    cropped = RotateBitmap(cropped, 
                                        RotateFlipType.Rotate180FlipNone);
                                    break;
                                case 270:
                                    cropped = RotateBitmap(cropped, 
                                        RotateFlipType.Rotate270FlipNone);
                                    break;
                            }
                        }

                        // contrast?
                        if (area.ContrastAdjustment != 0)
                        {
                            cropped = AdjustContrast(cropped, area.ContrastAdjustment);
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
```

If you walk through the code and read the comments, it should be easy to see how it all works at a high level.

Let's look at some of the internal functions. You can call these yourself from outside the library

**LoadOrFindBrowser**:

This method looks for the Microsoft Edge window with the site name using the **FindProcessByWindowTitle** method. If it's not running, it starts the process, looks for it again using **FindProcessByWindowTitle**, moves it to the desired screen, and presses {F11} to run it in kiosk mode.

```c#
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
```

**BitmapsAreEqual**:

This function returns true if two bitmaps are equal. This is used by the demo to determine whether the screen is frozen.

```c#
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
```

**GetText**:

This method uses the **Tesseract** library to perform Optical Character Recognition (OCR) on a given Bitmap.

**Note:** Tesseract requires metadata for the language that you are trying to recognize. I have included the training data for the English language. More language files can be download at https://github.com/tesseract-ocr/langdata

```c#
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
```

**CropBitmap**:

This function returns a new bitmap cropped to a particular rectangle

```c#
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
```

**RotateBitmap**:

This function returns a new bitmap rotated 90, 180, or 270 degrees to the right

```c#
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
```

**AdjustContrast**:

Sometimes you might need to adjust the contrast of an area before reading the text in order to get a match. A positive number increases the contrast. A negative number decreases the contrast.

```c#
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
```

**Click**:

Click clicks the mouse at a given coordinate.

```c#
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
```

**FindHighestContrastRectangle**:

This function returns an area of a given size of a bitmap that has the highest contrast. This is not used in the demo, but can come in handy in certain scenarios to find text when you don't know where it is.

```c#
/// <summary>
/// Returns an area of a given size of a bitmap that has the highest contrast
/// </summary>
/// <param name="bitmap">The Bitmap to examine</param>
/// <param name="size">The size of the rectangle to examine and return</param>
/// <param name="highestStdDev">THe highest standard deviation allowed</param>
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
```

**ComputeStandardDeviation**:

Called by FindHighestContrastRectangle to compute standard deviation.

```c#
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
```

**ProcessUIAction**:

This method processes either a Click (with the hover feature) or a SendKeys action

```c#
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
```

