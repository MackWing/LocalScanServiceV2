namespace LocalScanServiceV2.Models
{
    public class ScannedImage
    {
        public int Index { get; set; }
        public string DataUrl { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string Format { get; set; }
        public long SizeBytes { get; set; }
    }
}