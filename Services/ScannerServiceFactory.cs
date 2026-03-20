using LocalScanServiceV2.Models;
using System;
using System.Threading.Tasks;

namespace LocalScanServiceV2.Services
{
    public class ScannerServiceFactory
    {
        private IScannerService _wiaScannerService;
        private IScannerService _twainScannerService;

        public ScannerServiceFactory()
        {
            try
            {
                _wiaScannerService = new ScannerService();
                Console.WriteLine("WIA扫描服务初始化成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"WIA扫描服务初始化失败: {ex.Message}");
                _wiaScannerService = null;
            }

            try
            {
                _twainScannerService = new TwainScannerService();
                Console.WriteLine("TWAIN扫描服务初始化成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"TWAIN扫描服务初始化失败: {ex.Message}");
                _twainScannerService = null;
            }
        }

        public async Task<ScannersResponse> GetScannersAsync()
        {
            Console.WriteLine("开始获取扫描仪列表...");
            
            // 首先尝试WIA接口
            if (_wiaScannerService != null)
            {
                try
                {
                    var wiaResponse = await _wiaScannerService.GetScannersAsync();
                    if (wiaResponse.Success && wiaResponse.Scanners != null && wiaResponse.Scanners.Length > 0)
                    {
                        Console.WriteLine("WIA接口获取扫描仪列表成功");
                        return wiaResponse;
                    }
                    Console.WriteLine("WIA接口获取扫描仪列表失败，尝试TWAIN接口");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WIA接口获取扫描仪列表异常: {ex.Message}");
                    Console.WriteLine("尝试TWAIN接口");
                }
            }

            // 尝试TWAIN接口
            if (_twainScannerService != null)
            {
                try
                {
                    var twainResponse = await _twainScannerService.GetScannersAsync();
                    if (twainResponse.Success && twainResponse.Scanners != null && twainResponse.Scanners.Length > 0)
                    {
                        Console.WriteLine("TWAIN接口获取扫描仪列表成功");
                        return twainResponse;
                    }
                    Console.WriteLine("TWAIN接口获取扫描仪列表失败");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"TWAIN接口获取扫描仪列表异常: {ex.Message}");
                }
            }

            // 所有接口都失败
            var errorResponse = new ScannersResponse
            {
                Success = false,
                ErrorCode = "NO_SCANNERS_FOUND",
                ErrorMessage = "无法获取扫描仪列表，请确保：\n1. 设备已正确连接\n2. 设备驱动程序已正确安装\n3. 至少有一个接口(WIA或TWAIN)可用"
            };
            return errorResponse;
        }

        public async Task<ScanResponse> ScanAsync(ScanRequest request, string interfaceType = "auto")
        {
            Console.WriteLine($"开始扫描 - 接口类型: {interfaceType}");
            Console.WriteLine($"WIA服务状态: {(_wiaScannerService != null ? "可用" : "不可用")}");
            Console.WriteLine($"TWAIN服务状态: {(_twainScannerService != null ? "可用" : "不可用")}");

            // 自动模式：先尝试WIA，失败后尝试TWAIN
            if (interfaceType == "auto")
            {
                // 尝试WIA接口
                if (_wiaScannerService != null)
                {
                    try
                    {
                        Console.WriteLine("自动模式：尝试WIA接口");
                        var wiaResponse = await _wiaScannerService.ScanAsync(request);
                        if (wiaResponse.Success)
                        {
                            Console.WriteLine("WIA接口扫描成功");
                            return wiaResponse;
                        }
                        Console.WriteLine("WIA接口扫描失败，尝试TWAIN接口");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"WIA接口扫描异常: {ex.Message}");
                        Console.WriteLine("尝试TWAIN接口");
                    }
                }

                // 尝试TWAIN接口
                if (_twainScannerService != null)
                {
                    try
                    {
                        Console.WriteLine("自动模式：尝试TWAIN接口");
                        // 创建一个新的请求对象，清除设备ID，让TWAIN自动选择设备
                        var twainRequest = new ScanRequest
                        {
                            ScannerId = null, // 清除设备ID，让TWAIN自动选择
                            Dpi = request.Dpi,
                            ColorMode = request.ColorMode,
                            Format = request.Format,
                            Quality = request.Quality,
                            PaperSource = request.PaperSource,
                            MultiPage = request.MultiPage,
                            Brightness = request.Brightness,
                            Contrast = request.Contrast,
                            InterfaceType = request.InterfaceType
                        };
                        var twainResponse = await _twainScannerService.ScanAsync(twainRequest);
                        if (twainResponse.Success)
                        {
                            Console.WriteLine("TWAIN接口扫描成功");
                            return twainResponse;
                        }
                        Console.WriteLine("TWAIN接口扫描失败");
                        return twainResponse;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"TWAIN接口扫描异常: {ex.Message}");
                        var twainErrorResponse = new ScanResponse
                        {
                            Success = false,
                            ErrorCode = "SCAN_FAILED",
                            ErrorMessage = $"扫描失败: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保设备已正确连接\n2. 确保有文档放置在扫描区域\n3. 确保设备驱动程序已正确安装\n4. 确保设备没有被其他程序占用\n5. 尝试重启扫描仪和计算机\n6. 检查科达扫描仪的TWAIN驱动程序是否最新\n7. 对于科达Kodak i2600扫描仪，建议使用厂商提供的扫描软件"
                        };
                        return twainErrorResponse;
                    }
                }

                // 所有接口都失败
                var errorResponse = new ScanResponse
                {
                    Success = false,
                    ErrorCode = "SCAN_FAILED",
                    ErrorMessage = "所有扫描接口都失败，请确保：\n1. 设备已正确连接\n2. 有文档放置在扫描区域\n3. 设备驱动程序已正确安装\n4. 设备没有被其他程序占用\n5. 尝试重启扫描仪和计算机\n6. 检查科达扫描仪的驱动程序是否最新"
                };
                return errorResponse;
            }
            // 手动指定WIA接口
            else if (interfaceType == "wia" && _wiaScannerService != null)
            {
                try
                {
                    Console.WriteLine("手动指定：使用WIA接口");
                    var wiaResponse = await _wiaScannerService.ScanAsync(request);
                    if (wiaResponse.Success)
                    {
                        Console.WriteLine("WIA接口扫描成功");
                        return wiaResponse;
                    }
                    Console.WriteLine("WIA接口扫描失败");
                    return wiaResponse;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"WIA接口扫描异常: {ex.Message}");
                    var errorResponse = new ScanResponse
                    {
                        Success = false,
                        ErrorCode = "SCAN_FAILED",
                        ErrorMessage = $"WIA扫描失败: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保设备已正确连接\n2. 确保有文档放置在扫描区域\n3. 确保设备驱动程序已正确安装\n4. 确保设备没有被其他程序占用\n5. 尝试重启扫描仪和计算机\n6. 检查科达扫描仪的WIA驱动程序是否最新"
                    };
                    return errorResponse;
                }
            }
            // 手动指定TWAIN接口
            else if (interfaceType == "twain" && _twainScannerService != null)
            {
                try
                {
                    Console.WriteLine("手动指定：使用TWAIN接口");
                    // 创建一个新的请求对象，清除设备ID，让TWAIN自动选择设备
                    var twainRequest = new ScanRequest
                    {
                        ScannerId = null, // 清除设备ID，让TWAIN自动选择
                        Dpi = request.Dpi,
                        ColorMode = request.ColorMode,
                        Format = request.Format,
                        Quality = request.Quality,
                        PaperSource = request.PaperSource,
                        MultiPage = request.MultiPage,
                        Brightness = request.Brightness,
                        Contrast = request.Contrast,
                        InterfaceType = request.InterfaceType
                    };
                    var twainResponse = await _twainScannerService.ScanAsync(twainRequest);
                    if (twainResponse.Success)
                    {
                        Console.WriteLine("TWAIN接口扫描成功");
                        return twainResponse;
                    }
                    Console.WriteLine("TWAIN接口扫描失败");
                    return twainResponse;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"TWAIN接口扫描异常: {ex.Message}");
                    var errorResponse = new ScanResponse
                    {
                        Success = false,
                        ErrorCode = "SCAN_FAILED",
                        ErrorMessage = $"TWAIN扫描失败: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保设备已正确连接\n2. 确保有文档放置在扫描区域\n3. 确保设备驱动程序已正确安装\n4. 确保设备没有被其他程序占用\n5. 尝试重启扫描仪和计算机\n6. 检查科达扫描仪的TWAIN驱动程序是否最新"
                    };
                    return errorResponse;
                }
            }
            else
            {
                Console.WriteLine($"指定的接口类型 {interfaceType} 不可用");
                var errorResponse = new ScanResponse
                {
                    Success = false,
                    ErrorCode = "INTERFACE_NOT_AVAILABLE",
                    ErrorMessage = "指定的扫描接口不可用"
                };
                return errorResponse;
            }
        }

        public bool IsWiaAvailable => _wiaScannerService != null;
        public bool IsTwainAvailable => _twainScannerService != null;
    }
}