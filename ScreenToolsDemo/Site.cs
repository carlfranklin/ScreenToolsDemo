using System.Collections.Generic;

public class Site
{
    public string Name { get; set; }
    public string Url { get; set; }
    public string WindowTitle { get; set; }
    public string Description { get; set; }
    public List<TextArea> TextAreas { get; set; } = new List<TextArea>();
}
