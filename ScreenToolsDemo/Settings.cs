using System.Collections.Generic;

public class Settings
{
    public int ScreenIndex { get; set; }
    public int IntervalInSeconds { get; set; } = 5;
    public List<Site> Sites { get; set; } = new List<Site>();
}
