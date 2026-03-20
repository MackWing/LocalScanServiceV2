using LocalScanServiceV2.Models;
using LocalScanServiceV2.Services;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;

namespace LocalScanServiceV2.Controllers
{
    [Route("api")]
    public class ScanController : Controller
    {
        private static readonly ScannerServiceFactory _scannerServiceFactory = new ScannerServiceFactory();

        public ScanController()
        {
        }

        [HttpGet("health")]
        public IActionResult HealthCheck()
        {
            return Ok(new
            {
                status = "ok",
                version = "1.0.0",
                timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                wiaAvailable = _scannerServiceFactory.IsWiaAvailable,
                twainAvailable = _scannerServiceFactory.IsTwainAvailable
            });
        }

        [HttpGet("scanners")]
        public async Task<IActionResult> GetScanners()
        {
            var response = await _scannerServiceFactory.GetScannersAsync();
            return Ok(response);
        }

        [HttpPost("scan")]
        public async Task<IActionResult> Scan([FromBody] ScanRequest request)
        {
            // 从请求中获取接口类型，默认为auto
            string interfaceType = request.InterfaceType ?? "auto";
            var response = await _scannerServiceFactory.ScanAsync(request, interfaceType);
            return Ok(response);
        }

        // 暂时注释掉这些方法，因为TWAIN实现中还没有实现
        /*
        [HttpPost("preview")]
        public async Task<IActionResult> Preview([FromBody] dynamic request)
        {
            string scannerId = request?.scannerId;
            var response = await _scannerService.PreviewAsync(scannerId);
            return Ok(response);
        }

        [HttpPost("scan/cancel")]
        public async Task<IActionResult> CancelScan()
        {
            var success = await _scannerService.CancelScanAsync();
            return Ok(new
            {
                success = success,
                message = "扫描已取消"
            });
        }

        [HttpGet("scanners/{scannerId}/capabilities")]
        public async Task<IActionResult> GetScannerCapabilities(string scannerId)
        {
            var response = await _scannerService.GetScannerCapabilitiesAsync(scannerId);
            return Ok(response);
        }
        */
    }
}