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