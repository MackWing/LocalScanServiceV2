namespace LocalScanServiceV2.Models
{
    public class ScanRequest
    {
        public string ScannerId { get; set; }
        public int Dpi { get; set; } = 300;
        public string ColorMode { get; set; } = "color";
        public string Format { get; set; } = "png";
        public int Quality { get; set; } = 90;
        public string PaperSource { get; set; } = "flatbed";
        public bool MultiPage { get; set; } = false;
        public int Brightness { get; set; } = 0;
        public int Contrast { get; set; } = 0;
        public string InterfaceType { get; set; } = "auto";
        public ScanArea Area { get; set; }
    }

    public class ScanArea
    {
        public int Left { get; set; } = 0;
        public int Top { get; set; } = 0;
        public int Width { get; set; } = 210;
        public int Height { get; set; } = 297;
    }
}