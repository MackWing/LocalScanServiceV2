namespace LocalScanServiceV2.Models
{
    public class ScannerInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Manufacturer { get; set; }
        public bool IsDefault { get; set; }
        public bool IsReady { get; set; }
    }

    public class ScannersResponse
    {
        public bool Success { get; set; }
        public ScannerInfo[] Scanners { get; set; }
        public string DefaultScannerId { get; set; }
        public string ErrorCode { get; set; }
        public string ErrorMessage { get; set; }
    }
}