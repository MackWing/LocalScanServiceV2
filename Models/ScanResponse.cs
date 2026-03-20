namespace LocalScanServiceV2.Models
{
    public class ScanResponse
    {
        public bool Success { get; set; }
        public double ScanTime { get; set; }
        public ScannedImage[] Images { get; set; }
        public int TotalPages { get; set; }
        public string ErrorCode { get; set; }
        public string ErrorMessage { get; set; }
        public string Timestamp { get; set; }
    }
}