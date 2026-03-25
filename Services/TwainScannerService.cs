using LocalScanServiceV2.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace LocalScanServiceV2.Services
{
    // 实现IWindowsMessageHook接口，用于TWAIN消息处理
    public class SimpleWindowsMessageHook : TwainDotNet.IWindowsMessageHook
    {
        public IntPtr WindowHandle { get; private set; }
        public bool UseFilter { get; set; }
        public TwainDotNet.FilterMessage FilterMessageCallback { get; set; }

        public SimpleWindowsMessageHook()
        {
            // 创建一个隐藏窗口用于处理TWAIN消息
            CreateHiddenWindow();
        }

        private void CreateHiddenWindow()
        {
            try
            {
                // 使用Windows API创建一个隐藏窗口
                var className = "TwainMessageWindow" + Guid.NewGuid().ToString();
                var hInstance = System.Runtime.InteropServices.Marshal.GetHINSTANCE(typeof(SimpleWindowsMessageHook).Module);

                // 保存委托引用，防止垃圾回收
                _wndProcDelegate = WndProc;

                // 注册窗口类
                var wndClass = new WNDCLASS
                {
                    lpfnWndProc = _wndProcDelegate,
                    hInstance = hInstance,
                    lpszClassName = className
                };

                // 注册窗口类
                var atom = RegisterClass(ref wndClass);
                if (atom == 0)
                {
                    Console.WriteLine("注册窗口类失败");
                    return;
                }

                // 创建隐藏窗口
                WindowHandle = CreateWindowEx(0, className, "", 0, 0, 0, 0, 0, IntPtr.Zero, IntPtr.Zero, hInstance, IntPtr.Zero);
                if (WindowHandle == IntPtr.Zero)
                {
                    Console.WriteLine("创建窗口失败");
                }
                else
                {
                    Console.WriteLine("创建TWAIN消息窗口成功");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建隐藏窗口失败: {ex.Message}");
            }
        }

        public void Enable()
        {
            // 启用消息钩子
            UseFilter = true;
            Console.WriteLine("启用TWAIN消息钩子");
        }

        public void Disable()
        {
            // 禁用消息钩子
            UseFilter = false;
            Console.WriteLine("禁用TWAIN消息钩子");
        }

        // 窗口过程函数
        private IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
        {
            // 处理TWAIN消息
            return DefWindowProc(hWnd, msg, wParam, lParam);
        }

        // 保存委托引用，防止垃圾回收
        private WndProcDelegate _wndProcDelegate;

        // Windows API 声明
        private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct WNDCLASS
        {
            public uint style;
            public WndProcDelegate lpfnWndProc;
            public int cbClsExtra;
            public int cbWndExtra;
            public IntPtr hInstance;
            public IntPtr hIcon;
            public IntPtr hCursor;
            public IntPtr hbrBackground;
            public string lpszMenuName;
            public string lpszClassName;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern ushort RegisterClass(ref WNDCLASS lpWndClass);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CreateWindowEx(uint dwExStyle, string lpClassName, string lpWindowName, uint dwStyle, int x, int y, int nWidth, int nHeight, IntPtr hWndParent, IntPtr hMenu, IntPtr hInstance, IntPtr lpParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr DefWindowProc(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam);
    }

    public class TwainScannerService : IScannerService
    {
        private TwainDotNet.Twain _twain;
        private SimpleWindowsMessageHook _messageHook;

        public TwainScannerService()
        {
            try
            {
                // 创建消息钩子
                _messageHook = new SimpleWindowsMessageHook();
                Console.WriteLine("创建TWAIN消息钩子成功");
                
                // 初始化TWAIN
                _twain = new TwainDotNet.Twain(_messageHook);
                Console.WriteLine("TWAIN接口初始化成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"TWAIN接口初始化失败: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                _twain = null;
            }
        }

        public async Task<ScannersResponse> GetScannersAsync()
        {
            var response = new ScannersResponse();
            try
            {
                if (_twain == null)
                {
                    response.Success = false;
                    response.ErrorCode = "TWAIN_NOT_AVAILABLE";
                    response.ErrorMessage = "TWAIN接口不可用";
                    return response;
                }

                Console.WriteLine("开始获取TWAIN设备列表...");
                
                try
                {
                    // 获取TWAIN设备列表
                    var sourceNames = _twain.SourceNames;
                    Console.WriteLine($"找到 {sourceNames.Count} 个TWAIN设备");

                    var scanners = new List<ScannerInfo>();
                    string defaultScannerId = null;

                    for (int i = 0; i < sourceNames.Count; i++)
                    {
                        var sourceName = sourceNames[i];
                        var scannerInfo = new ScannerInfo
                        {
                            Id = sourceName,
                            Name = sourceName,
                            Manufacturer = "TWAIN设备",
                            IsDefault = i == 0,
                            IsReady = true
                        };
                        scanners.Add(scannerInfo);
                        Console.WriteLine($"添加TWAIN设备: {scannerInfo.Name}");

                        if (defaultScannerId == null)
                        {
                            defaultScannerId = scannerInfo.Id;
                        }
                    }

                    response.Success = true;
                    response.Scanners = scanners.ToArray();
                    response.DefaultScannerId = defaultScannerId;
                    Console.WriteLine($"TWAIN设备列表获取完成，找到 {scanners.Count} 个扫描仪");
                }
                catch (Exception ex)
                {
                    response.Success = false;
                    response.ErrorCode = "TWAIN_DEVICE_ENUMERATION_FAILED";
                    response.ErrorMessage = $"枚举TWAIN设备失败: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保TWAIN驱动程序已正确安装\n2. 尝试重启扫描仪和计算机\n3. 检查扫描仪是否支持TWAIN接口";
                    Console.WriteLine($"枚举TWAIN设备异常: {ex.Message}");
                    return response;
                }
            }
            catch (Exception ex)
            {
                response.Success = false;
                response.ErrorCode = "INTERNAL_ERROR";
                response.ErrorMessage = $"获取TWAIN扫描仪列表失败: {ex.Message}";
                Console.WriteLine($"获取TWAIN扫描仪列表异常: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
            return response;
        }

        public async Task<ScanResponse> ScanAsync(ScanRequest request)
        {
            var response = new ScanResponse();
            try
            {
                if (_twain == null)
                {
                    response.Success = false;
                    response.ErrorCode = "TWAIN_NOT_AVAILABLE";
                    response.ErrorMessage = "TWAIN接口不可用";
                    return response;
                }

                Console.WriteLine($"开始TWAIN扫描 - 设备ID: {request.ScannerId}");

                // 选择TWAIN设备
                try
                {
                    if (!string.IsNullOrEmpty(request.ScannerId))
                    {
                        // 尝试获取所有可用的TWAIN设备
                        var sourceNames = _twain.SourceNames;
                        Console.WriteLine($"可用的TWAIN设备数量: {sourceNames.Count}");
                        
                        // 检查设备ID是否在可用设备列表中
                        bool deviceFound = false;
                        foreach (var sourceName in sourceNames)
                        {
                            Console.WriteLine($"可用设备: {sourceName}");
                            if (sourceName.Equals(request.ScannerId, StringComparison.OrdinalIgnoreCase))
                            {
                                deviceFound = true;
                                break;
                            }
                        }
                        
                        if (deviceFound)
                        {
                            _twain.SelectSource(request.ScannerId);
                            Console.WriteLine($"选择TWAIN设备成功: {request.ScannerId}");
                        }
                        else
                        {
                            // 如果指定的设备ID不存在，使用默认设备
                            Console.WriteLine($"指定的设备ID {request.ScannerId} 不存在，使用默认设备");
                            _twain.SelectSource();
                            Console.WriteLine("选择默认TWAIN设备成功");
                        }
                    }
                    else
                    {
                        _twain.SelectSource();
                        Console.WriteLine("选择默认TWAIN设备成功");
                    }
                }
                catch (Exception ex)
                {
                    response.Success = false;
                    response.ErrorCode = "DEVICE_SELECTION_FAILED";
                    response.ErrorMessage = $"选择TWAIN设备失败: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保设备已正确连接\n2. 确保TWAIN驱动程序已正确安装\n3. 尝试重启扫描仪和计算机";
                    Console.WriteLine($"选择TWAIN设备异常: {ex.Message}");
                    return response;
                }

                // 执行扫描
                Console.WriteLine("开始执行TWAIN扫描...");
                var images = new List<Bitmap>();
                
                // 先移除之前的事件处理器，避免重复注册
                _twain.TransferImage -= (sender, e) => { };
                
                // 处理TWAIN扫描完成事件
                EventHandler<TwainDotNet.TransferImageEventArgs> transferImageHandler = (sender, e) =>
                {
                    Console.WriteLine("TransferImage事件触发...");
                    if (e != null)
                    {
                        Console.WriteLine("TransferImage事件参数不为空...");
                        if (e.Image != null)
                        {
                            Console.WriteLine($"获取到一张图像: {e.Image.Width}x{e.Image.Height}");
                            images.Add(e.Image);
                            Console.WriteLine($"图像已添加到列表，当前数量: {images.Count}");
                        }
                        else
                        {
                            Console.WriteLine("TransferImage事件参数的Image为空...");
                        }
                    }
                    else
                    {
                        Console.WriteLine("TransferImage事件参数为空...");
                    }
                };
                
                Console.WriteLine("注册TransferImage事件处理器...");
                _twain.TransferImage += transferImageHandler;
                Console.WriteLine("TransferImage事件处理器注册完成...");

                try
                {
                    // 开始扫描
                    var scanSettings = new TwainDotNet.ScanSettings();
                    // 显示TWAIN UI，让用户选择设备和设置
                    scanSettings.ShowTwainUI = true;
                    // 尝试设置一些基本参数
                    scanSettings.UseDocumentFeeder = true;
                    scanSettings.ShouldTransferAllPages = true;
                    
                    Console.WriteLine("准备开始TWAIN扫描...");
                    Console.WriteLine("TWAIN扫描设置: ShowTwainUI={0}, UseDocumentFeeder={1}, ShouldTransferAllPages={2}", 
                        scanSettings.ShowTwainUI, scanSettings.UseDocumentFeeder, scanSettings.ShouldTransferAllPages);
                    
                    // 创建一个事件来等待扫描完成
                    var scanCompleteEvent = new System.Threading.ManualResetEvent(false);
                    
                    // 启动扫描在一个新线程中，以便UI能够响应
                    System.Threading.Thread scanThread = new System.Threading.Thread(() =>
                    {
                        try
                        {
                            Console.WriteLine("TWAIN扫描线程启动...");
                            _twain.StartScanning(scanSettings);
                            Console.WriteLine("TWAIN扫描线程完成...");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"TWAIN扫描线程异常: {ex.Message}");
                        }
                        finally
                        {
                            // 通知主线程扫描已完成
                            scanCompleteEvent.Set();
                            Console.WriteLine("TWAIN扫描完成事件已设置...");
                        }
                    });
                    
                    // 启动扫描线程
                    scanThread.Start();
                    Console.WriteLine("TWAIN扫描线程已启动...");
                    
                    // 等待扫描完成，最多等待60秒
                    Console.WriteLine("等待TWAIN扫描完成...");
                    if (scanCompleteEvent.WaitOne(60000))
                    {
                        Console.WriteLine("TWAIN扫描完成，等待事件处理...");
                        // 等待一小段时间，确保所有事件都已处理完成
                        System.Threading.Thread.Sleep(2000);
                        
                        // 检查是否有图像获取到
                        if (images.Count > 0)
                        {
                            Console.WriteLine("TWAIN扫描成功获取到图像");
                            Console.WriteLine($"共获取到 {images.Count} 张图像");
                        }
                        else
                        {
                            Console.WriteLine("TWAIN扫描完成但未获取到图像");
                            Console.WriteLine($"共获取到 {images.Count} 张图像");
                        }
                    }
                    else
                    {
                        Console.WriteLine("TWAIN扫描超时");
                    }
                }
                catch (TwainDotNet.TwainException ex)
                {
                    // 移除事件处理器
                    _twain.TransferImage -= transferImageHandler;
                    
                    response.Success = false;
                    response.ErrorCode = "TWAIN_ERROR";
                    response.ErrorMessage = $"TWAIN扫描错误: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保扫描仪已正确连接并开启\n2. 确保TWAIN驱动程序已正确安装\n3. 确保扫描仪没有被其他程序占用\n4. 尝试重启扫描仪和计算机\n5. 检查科达扫描仪的TWAIN驱动程序是否最新\n6. 尝试使用厂商提供的扫描软件";
                    Console.WriteLine($"TWAIN扫描错误: {ex.Message}");
                    Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                    return response;
                }
                catch (Exception ex)
                {
                    // 移除事件处理器
                    _twain.TransferImage -= transferImageHandler;
                    
                    if (ex.Message.Contains("cancel"))
                    {
                        response.Success = false;
                        response.ErrorCode = "SCAN_CANCELED";
                        response.ErrorMessage = "扫描已被用户取消";
                        Console.WriteLine("扫描已被用户取消");
                        return response;
                    }
                    // 捕获空引用异常并提供更详细的错误信息
                    if (ex is NullReferenceException)
                    {
                        response.Success = false;
                        response.ErrorCode = "TWAIN_NULL_REFERENCE";
                        response.ErrorMessage = $"TWAIN扫描空引用异常: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保TWAIN驱动程序已正确安装\n2. 尝试重启扫描仪和计算机\n3. 检查扫描仪是否支持TWAIN接口";
                        Console.WriteLine($"TWAIN扫描空引用异常: {ex.Message}");
                        Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                        return response;
                    }
                    throw;
                }
                finally
                {
                    // 确保移除事件处理器
                    _twain.TransferImage -= transferImageHandler;
                }

                if (images.Count == 0)
                {
                    response.Success = false;
                    response.ErrorCode = "NO_IMAGES_SCANNED";
                    response.ErrorMessage = "扫描完成但未获取到图像\n\n请尝试以下解决方案：\n1. 确保有文档放置在扫描区域\n2. 检查扫描仪的进纸器是否正常\n3. 尝试手动放置文档\n4. 检查TWAIN驱动程序是否正确安装\n5. 尝试重启扫描仪和计算机";
                    Console.WriteLine("TWAIN扫描完成但未获取到图像");
                    return response;
                }

                // 处理扫描结果
                var scannedImages = new List<ScannedImage>();
                try
                {
                    for (int i = 0; i < images.Count; i++)
                    {
                        var image = images[i];
                        Console.WriteLine($"处理第 {i + 1} 张图像...");

                        // 将图像转换为Base64
                        string base64Image;
                        using (var ms = new MemoryStream())
                        {
                            image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            byte[] imageBytes = ms.ToArray();
                            base64Image = Convert.ToBase64String(imageBytes);
                        }

                        var scannedImage = new ScannedImage
                        {
                            Index = i,
                            DataUrl = "data:image/jpeg;base64," + base64Image,
                            Format = "jpeg",
                            Width = image.Width,
                            Height = image.Height,
                            SizeBytes = image.Width * image.Height * 3 // 估算大小
                        };
                        scannedImages.Add(scannedImage);
                        Console.WriteLine($"第 {i + 1} 张图像处理完成");
                    }
                }
                catch (Exception ex)
                {
                    response.Success = false;
                    response.ErrorCode = "IMAGE_PROCESSING_FAILED";
                    response.ErrorMessage = $"图像处理失败: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保有足够的内存\n2. 尝试降低扫描分辨率\n3. 检查图像格式是否支持";
                    Console.WriteLine($"图像处理异常: {ex.Message}");
                    return response;
                }

                response.Success = true;
                response.Images = scannedImages.ToArray();
                response.TotalPages = scannedImages.Count;
                response.Timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                Console.WriteLine("TWAIN扫描成功");
            }
            catch (Exception ex)
            {
                response.Success = false;
                response.ErrorCode = "SCAN_FAILED";
                response.ErrorMessage = $"TWAIN扫描失败: {ex.Message}\n\n请尝试以下解决方案：\n1. 确保设备已正确连接\n2. 确保有文档放置在扫描区域\n3. 确保设备驱动程序已正确安装\n4. 确保设备没有被其他程序占用\n5. 尝试重启扫描仪和计算机\n6. 检查科达扫描仪的TWAIN驱动程序是否最新";
                Console.WriteLine($"TWAIN扫描异常: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
            return response;
        }

        public async Task<ScanResponse> PreviewAsync(string scannerId)
        {
            var response = new ScanResponse();
            response.Success = false;
            response.ErrorCode = "NOT_IMPLEMENTED";
            response.ErrorMessage = "TWAIN接口暂不支持预览功能";
            return response;
        }

        public async Task<object> GetScannerCapabilitiesAsync(string scannerId)
        {
            var response = new
            {
                success = false,
                error = "NOT_IMPLEMENTED",
                message = "TWAIN接口暂不支持获取设备能力"
            };
            return response;
        }

        public async Task<bool> CancelScanAsync()
        {
            // TWAIN接口暂不支持取消扫描
            return false;
        }
    }
}