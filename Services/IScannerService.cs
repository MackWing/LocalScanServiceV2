using LocalScanServiceV2.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace LocalScanServiceV2.Services
{
    public interface IScannerService
    {
        Task<ScannersResponse> GetScannersAsync();
        Task<ScanResponse> ScanAsync(ScanRequest request);
        Task<ScanResponse> PreviewAsync(string scannerId);
        Task<object> GetScannerCapabilitiesAsync(string scannerId);
        Task<bool> CancelScanAsync();
    }
}