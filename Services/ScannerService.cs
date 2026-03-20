using LocalScanServiceV2.Models;
using LocalScanServiceV2.WIA;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LocalScanServiceV2.Services
{
    public class ScannerService : IScannerService
    {
        private object _deviceManager;
        private Type _deviceManagerType;

        public ScannerService()
        {
            try
            {
                _deviceManagerType = Type.GetTypeFromProgID("WIA.DeviceManager");
                if (_deviceManagerType != null)
                {
                    _deviceManager = Activator.CreateInstance(_deviceManagerType);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"初始化WIA设备管理器失败: {ex.Message}");
            }
        }

        public async Task<ScannersResponse> GetScannersAsync()
        {
            var response = new ScannersResponse();
            try
            {
                if (_deviceManager == null)
                {
                    response.Success = false;
                    response.ErrorCode = "WIA_NOT_AVAILABLE";
                    response.ErrorMessage = "WIA设备管理器不可用";
                    return response;
                }

                Console.WriteLine("开始获取设备列表...");
                var deviceInfos = _deviceManagerType.InvokeMember("DeviceInfos", BindingFlags.GetProperty, null, _deviceManager, null);
                Console.WriteLine("获取DeviceInfos成功");
                
                var deviceInfosType = deviceInfos.GetType();
                Console.WriteLine($"DeviceInfos类型: {deviceInfosType.FullName}");
                
                var count = (int)deviceInfosType.InvokeMember("Count", BindingFlags.GetProperty, null, deviceInfos, null);
                Console.WriteLine($"设备数量: {count}");
                Console.WriteLine($"WIA_DEVICE_SCANNER常量值: {WiaConstants.WIA_DEVICE_SCANNER}");

                var scanners = new List<ScannerInfo>();
                string defaultScannerId = null;

                for (int i = 1; i <= count; i++)
                {
                    Console.WriteLine($"处理设备 #{i}");
                    try
                    {
                        // 尝试使用不同的绑定标志获取设备信息
                        var deviceInfo = deviceInfosType.InvokeMember("Item", 
                            BindingFlags.InvokeMethod | BindingFlags.GetProperty, 
                            null, deviceInfos, new object[] { i });
                        
                        if (deviceInfo != null)
                        {
                            var deviceInfoType = deviceInfo.GetType();
                            Console.WriteLine($"设备信息类型: {deviceInfoType.FullName}");
                            
                            // 尝试获取设备的所有可用属性
                            string deviceName = $"设备 {i}";
                            string deviceManufacturer = "未知";
                            string deviceId = $"device_{i}";
                            
                            try
                            {
                                Console.WriteLine("=== 尝试获取设备属性 ===");
                                
                                // 详细记录设备属性获取过程
                                Console.WriteLine("=== 开始详细获取设备属性 ===");
                                
                                // 尝试使用不同的属性名称获取设备信息
                                string[] nameProperties = { "Name", "DeviceName", "FriendlyName", "Description", "Caption", "ProductName", "ItemName", "DisplayName", "Device", "ScannerName", "FriendlyDeviceName", "Product", "Name", "DeviceName" };
                                string[] idProperties = { "DeviceID", "Id", "UniqueId", "HardwareID", "DeviceId", "UID", "ID", "UniqueID", "HardwareId" };
                                string[] manufacturerProperties = { "Manufacturer", "Vendor", "Company", "Maker", "ManufacturerName", "VendorName", "CompanyName" };
                                
                                // 尝试获取设备名称
                                Console.WriteLine("--- 尝试获取设备名称 ---");
                                foreach (var propName in nameProperties)
                                {
                                    try
                                    {
                                        var value = deviceInfoType.InvokeMember(propName, 
                                            BindingFlags.GetProperty, null, deviceInfo, null);
                                        Console.WriteLine($"尝试获取 {propName}: {value}");
                                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                        {
                                            deviceName = value.ToString();
                                            Console.WriteLine($"✓ 找到设备名称 ({propName}): {deviceName}");
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"✗ 获取{propName}失败: {ex.Message}");
                                    }
                                }
                                
                                // 尝试获取设备ID
                                Console.WriteLine("--- 尝试获取设备ID ---");
                                foreach (var propName in idProperties)
                                {
                                    try
                                    {
                                        var value = deviceInfoType.InvokeMember(propName, 
                                            BindingFlags.GetProperty, null, deviceInfo, null);
                                        Console.WriteLine($"尝试获取 {propName}: {value}");
                                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                        {
                                            deviceId = value.ToString();
                                            Console.WriteLine($"✓ 找到设备ID ({propName}): {deviceId}");
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"✗ 获取{propName}失败: {ex.Message}");
                                    }
                                }
                                
                                // 尝试获取制造商信息
                                Console.WriteLine("--- 尝试获取制造商信息 ---");
                                foreach (var propName in manufacturerProperties)
                                {
                                    try
                                    {
                                        var value = deviceInfoType.InvokeMember(propName, 
                                            BindingFlags.GetProperty, null, deviceInfo, null);
                                        Console.WriteLine($"尝试获取 {propName}: {value}");
                                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                        {
                                            deviceManufacturer = value.ToString();
                                            Console.WriteLine($"✓ 找到制造商信息 ({propName}): {deviceManufacturer}");
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"✗ 获取{propName}失败: {ex.Message}");
                                    }
                                }
                                
                                // 尝试使用WIA特定的属性ID获取信息
                                try
                                {
                                    Console.WriteLine("=== 尝试使用WIA属性ID获取信息 ===");
                                    // 尝试获取设备的Properties集合
                                    var properties = deviceInfoType.InvokeMember("Properties", BindingFlags.GetProperty, null, deviceInfo, null);
                                    if (properties != null)
                                    {
                                        var propertiesType = properties.GetType();
                                        var propertiesCount = (int)propertiesType.InvokeMember("Count", BindingFlags.GetProperty, null, properties, null);
                                        Console.WriteLine($"属性数量: {propertiesCount}");
                                        
                                        // 尝试遍历所有属性，查找可能的设备名称属性
                                        for (int j = 1; j <= propertiesCount; j++)
                                        {
                                            try
                                            {
                                                var property = propertiesType.InvokeMember("Item", BindingFlags.InvokeMethod, null, properties, new object[] { j });
                                                if (property != null)
                                                {
                                                    var propertyType = property.GetType();
                                                    var propertyId = (int)propertyType.InvokeMember("PropertyID", BindingFlags.GetProperty, null, property, null);
                                                    var propertyName = propertyType.InvokeMember("Name", BindingFlags.GetProperty, null, property, null);
                                                    var propertyValue = propertyType.InvokeMember("Value", BindingFlags.GetProperty, null, property, null);
                                                    
                                                    Console.WriteLine($"属性 #{j}: ID={propertyId}, Name={propertyName}, Value={propertyValue}");
                                                    
                                                    // 尝试使用WIA_DIP_DEVICENAME属性
                                                    if (propertyId == 3073) // WIA_DIP_DEVICENAME
                                                    {
                                                        if (propertyValue != null && !string.IsNullOrEmpty(propertyValue.ToString()))
                                                        {
                                                            deviceName = propertyValue.ToString();
                                                            Console.WriteLine($"✓ 找到设备名称 (WIA_DIP_DEVICENAME): {deviceName}");
                                                        }
                                                    }
                                                    // 尝试使用WIA_DIP_VENDOR属性
                                                    else if (propertyId == 3072) // WIA_DIP_VENDOR
                                                    {
                                                        if (propertyValue != null && !string.IsNullOrEmpty(propertyValue.ToString()))
                                                        {
                                                            deviceManufacturer = propertyValue.ToString();
                                                            Console.WriteLine($"✓ 找到制造商信息 (WIA_DIP_VENDOR): {deviceManufacturer}");
                                                        }
                                                    }
                                                    // 尝试使用其他可能的属性
                                                    else if (propertyName != null)
                                                    {
                                                        string propNameStr = propertyName.ToString().ToLower();
                                                        if ((propNameStr.Contains("name") || propNameStr.Contains("device")) && 
                                                            propertyValue != null && !string.IsNullOrEmpty(propertyValue.ToString()))
                                                        {
                                                            deviceName = propertyValue.ToString();
                                                            Console.WriteLine($"✓ 找到设备名称 (属性名包含name/device): {deviceName}");
                                                        }
                                                        else if (propNameStr.Contains("manufacturer") || propNameStr.Contains("vendor"))
                                                        {
                                                            if (propertyValue != null && !string.IsNullOrEmpty(propertyValue.ToString()))
                                                            {
                                                                deviceManufacturer = propertyValue.ToString();
                                                                Console.WriteLine($"✓ 找到制造商信息 (属性名包含manufacturer/vendor): {deviceManufacturer}");
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"✗ 处理属性 #{j} 失败: {ex.Message}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("✗ 获取Properties集合失败: 返回null");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"✗ 获取WIA属性失败: {ex.Message}");
                                }
                                
                                // 尝试使用不同的绑定标志获取属性
                                try
                                {
                                    Console.WriteLine("=== 尝试使用不同的绑定标志获取属性 ===");
                                    // 尝试使用不同的绑定标志组合
                                    BindingFlags[] bindingFlags = {
                                        BindingFlags.GetProperty,
                                        BindingFlags.GetProperty | BindingFlags.Instance,
                                        BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                                        BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public | BindingFlags.Static
                                    };
                                    
                                    foreach (var flags in bindingFlags)
                                    {
                                        try
                                        {
                                            var value = deviceInfoType.InvokeMember("Name", flags, null, deviceInfo, null);
                                            Console.WriteLine($"尝试使用绑定标志 {flags} 获取Name: {value}");
                                            if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                            {
                                                deviceName = value.ToString();
                                                Console.WriteLine($"✓ 找到设备名称 (使用绑定标志 {flags}): {deviceName}");
                                                break;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"✗ 使用绑定标志 {flags} 获取Name失败: {ex.Message}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"✗ 尝试不同绑定标志失败: {ex.Message}");
                                }
                                
                                // 尝试使用InvokeMethod获取属性
                                try
                                {
                                    Console.WriteLine("=== 尝试使用InvokeMethod获取属性 ===");
                                    var value = deviceInfoType.InvokeMember("Name", BindingFlags.InvokeMethod, null, deviceInfo, null);
                                    Console.WriteLine($"尝试使用InvokeMethod获取Name: {value}");
                                    if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                    {
                                        deviceName = value.ToString();
                                        Console.WriteLine($"✓ 找到设备名称 (使用InvokeMethod): {deviceName}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"✗ 使用InvokeMethod获取Name失败: {ex.Message}");
                                }
                                
                                // 打印最终获取的设备信息
                                Console.WriteLine("=== 最终设备信息 ===");
                                Console.WriteLine($"设备名称: {deviceName}");
                                Console.WriteLine($"设备ID: {deviceId}");
                                Console.WriteLine($"制造商: {deviceManufacturer}");
                                
                                // 尝试获取更多属性以帮助诊断
                                string[] otherProperties = { "Type", "Status", "Location", "DriverName" };
                                foreach (var propName in otherProperties)
                                {
                                    try
                                    {
                                        var value = deviceInfoType.InvokeMember(propName, 
                                            BindingFlags.GetProperty, null, deviceInfo, null);
                                        Console.WriteLine($"{propName}: {value}");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"获取{propName}失败: {ex.Message}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"获取设备属性失败: {ex.Message}");
                            }
                            
                            // 尝试使用不同的绑定标志
                            try
                            {
                                Console.WriteLine("=== 尝试不同的绑定标志 ===");
                                var type = (int)deviceInfoType.InvokeMember("Type", 
                                    BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public, 
                                    null, deviceInfo, null);
                                Console.WriteLine($"设备类型 (使用不同绑定标志): {type}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"获取设备类型失败: {ex.Message}");
                            }
                            
                            // 创建ScannerInfo对象
                            try
                            {
                                Console.WriteLine("=== 创建ScannerInfo ===");
                                var scannerInfo = new ScannerInfo
                                {
                                    Id = deviceId,
                                    Name = deviceName,
                                    Manufacturer = deviceManufacturer,
                                    IsDefault = false,
                                    IsReady = true
                                };
                                scanners.Add(scannerInfo);
                                Console.WriteLine($"添加设备: {scannerInfo.Name} (ID: {scannerInfo.Id}, 制造商: {scannerInfo.Manufacturer})");
                                
                                if (defaultScannerId == null)
                                {
                                    defaultScannerId = scannerInfo.Id;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"创建ScannerInfo失败: {ex.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"设备 #{i} 为null");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"处理设备 #{i} 时出错: {ex.Message}");
                        Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                        // 继续处理下一个设备
                    }
                }

                response.Success = true;
                response.Scanners = scanners.ToArray();
                response.DefaultScannerId = defaultScannerId;
                Console.WriteLine($"扫描完成，找到 {scanners.Count} 个扫描仪");
            }
            catch (Exception ex)
            {
                response.Success = false;
                response.ErrorCode = "INTERNAL_ERROR";
                response.ErrorMessage = $"获取扫描仪列表失败: {ex.Message}\n堆栈跟踪: {ex.StackTrace}";
                Console.WriteLine($"获取扫描仪列表异常: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
            return response;
        }

        public async Task<ScanResponse> ScanAsync(ScanRequest request)
        {
            var response = new ScanResponse();
            try
            {
                if (_deviceManager == null)
                {
                    response.Success = false;
                    response.ErrorCode = "WIA_NOT_AVAILABLE";
                    response.ErrorMessage = "WIA设备管理器不可用";
                    return response;
                }

                var startTime = DateTime.Now;
                Console.WriteLine($"开始扫描 - 设备ID: {request.ScannerId}, DPI: {request.Dpi}, 颜色模式: {request.ColorMode}");

                // 获取设备
                var device = GetDevice(request.ScannerId);
                if (device == null)
                {
                    response.Success = false;
                    response.ErrorCode = "SCANNER_NOT_FOUND";
                    response.ErrorMessage = "未找到指定的扫描仪";
                    return response;
                }
                Console.WriteLine("设备获取成功");

                var deviceType = device.GetType();
                Console.WriteLine($"设备类型: {deviceType.FullName}");

                // 设置扫描参数
                Console.WriteLine("开始设置扫描参数...");
                SetScanParameters(device, deviceType, request);
                Console.WriteLine("扫描参数设置完成");

                // 执行扫描
                var images = new List<ScannedImage>();
                try
                {
                    Console.WriteLine("获取Items...");
                    var items = deviceType.InvokeMember("Items", BindingFlags.GetProperty, null, device, null);
                    if (items == null)
                    {
                        Console.WriteLine("Items获取失败: 返回null");
                        throw new Exception("无法获取扫描项目，请确保设备已准备就绪且有文档放置在扫描区域");
                    }
                    Console.WriteLine("Items获取成功");
                    
                    var itemsType = items.GetType();
                    Console.WriteLine($"Items类型: {itemsType.FullName}");
                    
                    // 尝试获取Items的Count属性
                    try
                    {
                        var count = (int)itemsType.InvokeMember("Count", BindingFlags.GetProperty, null, items, null);
                        Console.WriteLine($"Items数量: {count}");
                        if (count == 0)
                        {
                            throw new Exception("设备没有可用的扫描项目，请确保设备已准备就绪");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"获取Items数量失败: {ex.Message}");
                        // 继续执行，即使无法获取Count
                    }
                    
                    Console.WriteLine("获取Item...");
                    object item = null;
                    
                    // 尝试不同的调用方式获取Item
                    try
                    {
                        // 方式1：使用InvokeMethod
                        item = itemsType.InvokeMember("Item", BindingFlags.InvokeMethod, null, items, new object[] { 1 });
                        Console.WriteLine("Item获取成功 (使用InvokeMethod)");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"使用InvokeMethod获取Item失败: {ex.Message}");
                        
                        try
                        {
                            // 方式2：使用GetProperty
                            item = itemsType.InvokeMember("Item", BindingFlags.GetProperty, null, items, new object[] { 1 });
                            Console.WriteLine("Item获取成功 (使用GetProperty)");
                        }
                        catch (Exception ex2)
                        {
                            Console.WriteLine($"使用GetProperty获取Item失败: {ex2.Message}");
                            
                            try
                            {
                                // 方式3：使用不同的绑定标志
                                item = itemsType.InvokeMember("Item", 
                                    BindingFlags.InvokeMethod | BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public, 
                                    null, items, new object[] { 1 });
                                Console.WriteLine("Item获取成功 (使用多种绑定标志)");
                            }
                            catch (Exception ex3)
                            {
                                Console.WriteLine($"所有方式获取Item都失败: {ex3.Message}");
                                throw new Exception($"无法获取扫描项目: {ex3.Message}\n请确保设备已准备就绪且有文档放置在扫描区域");
                            }
                        }
                    }
                    
                    if (item == null)
                    {
                        throw new Exception("获取到的Item为null，请确保设备已准备就绪");
                    }
                    
                    var itemType = item.GetType();
                    Console.WriteLine($"Item类型: {itemType.FullName}");

                    Console.WriteLine("开始传输图像...");
                    object imageFile = null;
                    
                    // 尝试获取设备状态
                    try
                    {
                        var statusProperty = deviceType.InvokeMember("Properties", BindingFlags.GetProperty, null, device, null);
                        if (statusProperty != null)
                        {
                            var statusPropertyType = statusProperty.GetType();
                            try
                            {
                                var status = statusPropertyType.InvokeMember("Item", BindingFlags.InvokeMethod, null, statusProperty, new object[] { 3076 }); // WIA_DIP_DEVICESTATUS
                                if (status != null)
                                {
                                    var statusType = status.GetType();
                                    var statusValue = statusType.InvokeMember("Value", BindingFlags.GetProperty, null, status, null);
                                    Console.WriteLine($"设备状态: {statusValue}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"获取设备状态失败: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"获取设备属性失败: {ex.Message}");
                    }
                    
                    // 尝试不同的Transfer调用方式，添加重试机制
                    int transferAttempts = 3;
                    for (int attempt = 1; attempt <= transferAttempts; attempt++)
                    {
                        try
                        {
                            Console.WriteLine($"尝试传输图像 (尝试 #{attempt}/{transferAttempts})...");
                            
                            // 记录设备和项目的详细信息
                            Console.WriteLine($"设备类型: {deviceType.FullName}");
                            Console.WriteLine($"项目类型: {itemType.FullName}");
                            
                            // 尝试获取设备的更多属性信息
                            try
                            {
                                var deviceProperties = deviceType.InvokeMember("Properties", BindingFlags.GetProperty, null, device, null);
                                if (deviceProperties != null)
                                {
                                    var devicePropertiesType = deviceProperties.GetType();
                                    var devicePropertiesCount = (int)devicePropertiesType.InvokeMember("Count", BindingFlags.GetProperty, null, deviceProperties, null);
                                    Console.WriteLine($"设备属性数量: {devicePropertiesCount}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"获取设备属性失败: {ex.Message}");
                            }
                            
                            // 方式1：使用InvokeMethod和不同的绑定标志，参数: 0
                            try
                            {
                                Console.WriteLine("尝试方式1: 使用InvokeMethod和完整绑定标志，参数: 0");
                                imageFile = itemType.InvokeMember("Transfer", 
                                    BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public, 
                                    null, item, new object[] { 0 });
                                Console.WriteLine("图像传输成功 (使用InvokeMethod和完整绑定标志)");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"方式1失败: {ex.Message}");
                                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                                
                                // 方式2：尝试使用不同的参数，参数: 2
                                try
                                {
                                    Console.WriteLine("尝试方式2: 使用InvokeMethod，参数: 2");
                                    imageFile = itemType.InvokeMember("Transfer", 
                                        BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public, 
                                        null, item, new object[] { 2 });
                                    Console.WriteLine("图像传输成功 (使用参数2)");
                                    break;
                                }
                                catch (Exception ex2)
                                {
                                    Console.WriteLine($"方式2失败: {ex2.Message}");
                                    
                                    // 方式3：尝试使用参数: 3
                                    try
                                    {
                                        Console.WriteLine("尝试方式3: 使用InvokeMethod，参数: 3");
                                        imageFile = itemType.InvokeMember("Transfer", 
                                            BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public, 
                                            null, item, new object[] { 3 });
                                        Console.WriteLine("图像传输成功 (使用参数3)");
                                        break;
                                    }
                                    catch (Exception ex3)
                                    {
                                        Console.WriteLine($"方式3失败: {ex3.Message}");
                                        
                                        // 方式4：尝试使用默认参数
                                        try
                                        {
                                            Console.WriteLine("尝试方式4: 使用默认参数");
                                            imageFile = itemType.InvokeMember("Transfer", BindingFlags.InvokeMethod, null, item, null);
                                            Console.WriteLine("图像传输成功 (使用默认参数)");
                                            break;
                                        }
                                        catch (Exception ex4)
                                        {
                                            Console.WriteLine($"方式4失败: {ex4.Message}");
                                            
                                            // 方式5：尝试使用GetProperty
                                            try
                                            {
                                                Console.WriteLine("尝试方式5: 使用GetProperty");
                                                imageFile = itemType.InvokeMember("Transfer", BindingFlags.GetProperty, null, item, null);
                                                Console.WriteLine("图像传输成功 (使用GetProperty)");
                                                break;
                                            }
                                            catch (Exception ex5)
                                            {
                                                Console.WriteLine($"方式5失败: {ex5.Message}");
                                                Console.WriteLine($"堆栈跟踪: {ex5.StackTrace}");
                                                
                                                // 方式6：尝试使用不同的绑定标志组合
                                                try
                                                {
                                                    Console.WriteLine("尝试方式6: 使用不同的绑定标志组合");
                                                    imageFile = itemType.InvokeMember("Transfer", 
                                                        BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public | BindingFlags.Static, 
                                                        null, item, null);
                                                    Console.WriteLine("图像传输成功 (使用多种绑定标志)");
                                                    break;
                                                }
                                                catch (Exception ex6)
                                                {
                                                    Console.WriteLine($"方式6失败: {ex6.Message}");
                                                    
                                                    // 方式7：尝试使用其他可能的方法名称
                                                    try
                                                    {
                                                        Console.WriteLine("尝试方式7: 使用Scan方法");
                                                        imageFile = itemType.InvokeMember("Scan", BindingFlags.InvokeMethod, null, item, null);
                                                        Console.WriteLine("图像传输成功 (使用Scan方法)");
                                                        break;
                                                    }
                                                    catch (Exception ex7)
                                                    {
                                                        Console.WriteLine($"方式7失败: {ex7.Message}");
                                                        
                                                        // 方式8：尝试使用Acquire方法
                                                        try
                                                        {
                                                            Console.WriteLine("尝试方式8: 使用Acquire方法");
                                                            imageFile = itemType.InvokeMember("Acquire", BindingFlags.InvokeMethod, null, item, null);
                                                            Console.WriteLine("图像传输成功 (使用Acquire方法)");
                                                            break;
                                                        }
                                                        catch (Exception ex8)
                                                        {
                                                            Console.WriteLine($"方式8失败: {ex8.Message}");
                                                            
                                                            // 如果是最后一次尝试，抛出异常
                                                            if (attempt == transferAttempts)
                                                            {
                                                                throw new Exception($"无法传输图像: {ex8.Message}\n请确保：\n1. 设备已正确连接\n2. 有文档放置在扫描区域\n3. 设备驱动程序已正确安装\n4. 设备没有被其他程序占用\n5. 尝试重启扫描仪和计算机\n6. 检查科达扫描仪的WIA驱动程序是否最新\n7. 尝试使用TWAIN接口替代WIA接口");
                                                            }
                                                            
                                                            // 等待一段时间后重试
                                                            Console.WriteLine("等待2秒后重试...");
                                                            System.Threading.Thread.Sleep(2000);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"传输尝试失败 (尝试 #{attempt}): {ex.Message}");
                            Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                            
                            // 如果是最后一次尝试，抛出异常
                            if (attempt == transferAttempts)
                            {
                                throw;
                            }
                            
                            // 等待一段时间后重试
                            Console.WriteLine("等待2秒后重试...");
                            System.Threading.Thread.Sleep(2000);
                        }
                    }
                    
                    if (imageFile == null)
                    {
                        throw new Exception("获取到的图像文件为null，请确保设备已准备就绪");
                    }
                    
                    var imageFileType = imageFile.GetType();
                    Console.WriteLine($"图像文件类型: {imageFileType.FullName}");

                    // 尝试直接获取图像数据，不使用SaveFile方法
                    try
                    {
                        Console.WriteLine("尝试直接获取图像数据...");
                        
                        // 尝试获取图像的FileData属性
                        object fileData = null;
                        try
                        {
                            fileData = imageFileType.InvokeMember("FileData", BindingFlags.GetProperty, null, imageFile, null);
                            Console.WriteLine("成功获取FileData属性");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"获取FileData属性失败: {ex.Message}");
                            Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                            
                            // 尝试其他可能的属性名称
                            try
                            {
                                Console.WriteLine("尝试使用Data属性获取图像数据...");
                                fileData = imageFileType.InvokeMember("Data", BindingFlags.GetProperty, null, imageFile, null);
                                Console.WriteLine("成功获取Data属性");
                            }
                            catch (Exception ex2)
                            {
                                Console.WriteLine($"获取Data属性失败: {ex2.Message}");
                                throw new Exception($"无法获取图像数据: {ex.Message}\n请确保设备驱动程序正确安装\n可能的原因: 科达扫描仪的WIA接口实现与其他品牌不同");
                            }
                        }
                        
                        if (fileData == null)
                        {
                            throw new Exception("FileData属性为null，请确保设备驱动程序正确安装");
                        }
                        
                        // 获取FileData的BinaryData
                        object binaryData = null;
                        try
                        {
                            var fileDataType = fileData.GetType();
                            binaryData = fileDataType.InvokeMember("BinaryData", BindingFlags.GetProperty, null, fileData, null);
                            Console.WriteLine("成功获取BinaryData属性");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"获取BinaryData属性失败: {ex.Message}");
                            Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                            
                            // 尝试其他可能的属性名称
                            try
                            {
                                Console.WriteLine("尝试使用Bytes属性获取图像数据...");
                                var fileDataType = fileData.GetType();
                                binaryData = fileDataType.InvokeMember("Bytes", BindingFlags.GetProperty, null, fileData, null);
                                Console.WriteLine("成功获取Bytes属性");
                            }
                            catch (Exception ex2)
                            {
                                Console.WriteLine($"获取Bytes属性失败: {ex2.Message}");
                                throw new Exception($"无法获取图像二进制数据: {ex.Message}\n请确保设备驱动程序正确安装");
                            }
                        }
                        
                        if (binaryData == null)
                        {
                            throw new Exception("BinaryData属性为null，请确保设备驱动程序正确安装");
                        }
                        
                        // 处理二进制数据
                        byte[] imageBytes = null;
                        if (binaryData is byte[])
                        {
                            imageBytes = (byte[])binaryData;
                            Console.WriteLine($"获取到图像数据，大小: {imageBytes.Length}字节");
                        }
                        else
                        {
                            Console.WriteLine($"二进制数据类型: {binaryData.GetType().FullName}");
                            
                            // 尝试将其他类型转换为字节数组
                            try
                            {
                                Console.WriteLine("尝试将二进制数据转换为字节数组...");
                                // 这里可以根据实际类型添加转换逻辑
                                throw new Exception($"不支持的二进制数据类型: {binaryData.GetType().FullName}");
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"无法处理图像数据: {ex.Message}");
                            }
                        }
                        
                        // 读取并处理图像
                        try
                        {
                            using (var ms = new MemoryStream(imageBytes))
                            {
                                using (var image = Image.FromStream(ms))
                                {
                                    using (var outputMs = new MemoryStream())
                                    {
                                        ImageFormat format = request.Format.ToLower() == "jpeg" ? ImageFormat.Jpeg : ImageFormat.Png;
                                        image.Save(outputMs, format);
                                        var bytes = outputMs.ToArray();
                                        var base64 = Convert.ToBase64String(bytes);
                                        var dataUrl = $"data:image/{request.Format.ToLower()};base64,{base64}";

                                        images.Add(new ScannedImage
                                        {
                                            Index = 0,
                                            DataUrl = dataUrl,
                                            Width = image.Width,
                                            Height = image.Height,
                                            Format = request.Format.ToLower(),
                                            SizeBytes = bytes.Length
                                        });
                                        Console.WriteLine($"图像处理完成 - 尺寸: {image.Width}x{image.Height}, 大小: {bytes.Length}字节");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"图像处理失败: {ex.Message}");
                            Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                            throw new Exception($"处理扫描图像失败: {ex.Message}\n可能的原因: 图像格式不兼容或损坏");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"处理图像数据失败: {ex.Message}");
                        Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                        throw new Exception($"处理扫描图像失败: {ex.Message}\n请确保：\n1. 设备驱动程序正确安装\n2. 图像格式设置正确\n3. 尝试使用不同的扫描参数");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"扫描执行失败: {ex.Message}");
                    Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                    throw;
                }

                var scanTime = (DateTime.Now - startTime).TotalSeconds;
                Console.WriteLine($"扫描完成 - 耗时: {scanTime}秒, 页数: {images.Count}");

                response.Success = true;
                response.ScanTime = scanTime;
                response.Images = images.ToArray();
                response.TotalPages = images.Count;
            }
            catch (Exception ex)
            {
                response.Success = false;
                response.ErrorCode = "SCAN_ERROR";
                response.ErrorMessage = $"扫描失败: {ex.Message}";
                response.Timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
                Console.WriteLine($"扫描异常: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
            return response;
        }

        public async Task<ScanResponse> PreviewAsync(string scannerId)
        {
            // 创建一个低分辨率的扫描请求
            var request = new ScanRequest
            {
                ScannerId = scannerId,
                Dpi = 100,
                ColorMode = "grayscale",
                Format = "jpeg",
                Quality = 70,
                PaperSource = "flatbed"
            };

            return await ScanAsync(request);
        }

        public async Task<object> GetScannerCapabilitiesAsync(string scannerId)
        {
            try
            {
                Console.WriteLine($"获取扫描仪能力 - 设备ID: {scannerId}");
                var device = GetDevice(scannerId);
                if (device == null)
                {
                    Console.WriteLine("未找到指定的扫描仪");
                    return new { success = false, errorCode = "SCANNER_NOT_FOUND", errorMessage = "未找到指定的扫描仪" };
                }
                Console.WriteLine("设备获取成功");

                var deviceType = device.GetType();
                Console.WriteLine($"设备类型: {deviceType.FullName}");
                
                object properties = null;
                try
                {
                    properties = deviceType.InvokeMember("Properties", BindingFlags.GetProperty, null, device, null);
                    if (properties == null)
                    {
                        Console.WriteLine("获取Properties失败: 返回null");
                        return new { success = false, errorCode = "PROPERTIES_NOT_FOUND", errorMessage = "无法获取设备属性" };
                    }
                    Console.WriteLine("Properties获取成功");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"获取Properties失败: {ex.Message}");
                    return new { success = false, errorCode = "PROPERTIES_ERROR", errorMessage = $"获取设备属性失败: {ex.Message}" };
                }
                
                var propertiesType = properties.GetType();
                Console.WriteLine($"Properties类型: {propertiesType.FullName}");
                
                int propertiesCount = 0;
                try
                {
                    propertiesCount = (int)propertiesType.InvokeMember("Count", BindingFlags.GetProperty, null, properties, null);
                    Console.WriteLine($"属性数量: {propertiesCount}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"获取属性数量失败: {ex.Message}");
                    // 继续执行，使用默认值
                }

                var capabilities = new
                {
                    maxDpi = 1200,
                    minDpi = 100,
                    supportedDpi = new int[] { 100, 150, 200, 300, 600, 1200 },
                    supportedColorModes = new string[] { "color", "grayscale", "blackwhite" },
                    supportedFormats = new string[] { "png", "jpeg", "bmp" },
                    hasFeeder = false,
                    hasDuplex = false,
                    maxScanArea = new { width = 216, height = 356 }
                };

                // 尝试从设备属性中获取实际能力
                if (propertiesCount > 0)
                {
                    for (int i = 1; i <= propertiesCount; i++)
                    {
                        try
                        {
                            var property = propertiesType.InvokeMember("Item", BindingFlags.InvokeMethod, null, properties, new object[] { i });
                            if (property != null)
                            {
                                var propertyType = property.GetType();
                                var propertyId = (int)propertyType.InvokeMember("PropertyID", BindingFlags.GetProperty, null, property, null);
                                Console.WriteLine($"属性 #{i}: ID={propertyId}");

                                // 这里可以根据需要解析更多属性
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"处理属性 #{i} 失败: {ex.Message}");
                            // 继续处理下一个属性
                        }
                    }
                }

                Console.WriteLine("获取扫描仪能力成功");
                return new { success = true, capabilities = capabilities };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取扫描仪能力异常: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                return new { success = false, errorCode = "INTERNAL_ERROR", errorMessage = $"获取扫描仪能力失败: {ex.Message}" };
            }
        }

        public async Task<bool> CancelScanAsync()
        {
            // WIA 不支持直接取消扫描，这里返回true表示取消操作已发出
            return true;
        }

        private object GetDevice(string scannerId)
        {
            try
            {
                var deviceInfos = _deviceManagerType.InvokeMember("DeviceInfos", BindingFlags.GetProperty, null, _deviceManager, null);
                var deviceInfosType = deviceInfos.GetType();
                var count = (int)deviceInfosType.InvokeMember("Count", BindingFlags.GetProperty, null, deviceInfos, null);

                Console.WriteLine($"GetDevice: scannerId={scannerId}, 设备数量={count}");

                // 检查是否是我们生成的设备ID格式 (device_1, device_2, etc.)
                int deviceIndex = -1;
                if (scannerId != null && scannerId.StartsWith("device_"))
                {
                    try
                    {
                        deviceIndex = int.Parse(scannerId.Substring(7));
                        Console.WriteLine($"识别到生成的设备ID，索引={deviceIndex}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"解析设备ID失败: {ex.Message}");
                    }
                }

                // 存储所有可用的扫描仪设备
                List<object> scannerDevices = new List<object>();

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        var deviceInfo = deviceInfosType.InvokeMember("Item", 
                            BindingFlags.InvokeMethod | BindingFlags.GetProperty, 
                            null, deviceInfos, new object[] { i });
                        
                        if (deviceInfo != null)
                        {
                            // 如果是我们生成的设备ID格式，直接按索引匹配
                            if (deviceIndex > 0 && i == deviceIndex)
                            {
                                Console.WriteLine($"按索引匹配设备 #{i}");
                                try
                                {
                                    var deviceInfoType = deviceInfo.GetType();
                                    var device = deviceInfoType.InvokeMember("Connect", 
                                        BindingFlags.InvokeMethod, null, deviceInfo, null);
                                    Console.WriteLine($"成功连接设备 #{i}");
                                    return device;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"连接设备 #{i} 失败: {ex.Message}");
                                    Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                                }
                            }
                            
                            // 尝试使用原始设备ID匹配
                            try
                            {
                                var deviceInfoType = deviceInfo.GetType();
                                var type = (int)deviceInfoType.InvokeMember("Type", 
                                    BindingFlags.GetProperty, null, deviceInfo, null);
                                
                                if (type == WiaConstants.WIA_DEVICE_SCANNER)
                                {
                                    var id = (string)deviceInfoType.InvokeMember("DeviceID", 
                                        BindingFlags.GetProperty, null, deviceInfo, null);
                                    Console.WriteLine($"设备 #{i} 类型: {type}, ID: {id}");
                                    
                                    // 存储扫描仪设备信息
                                    scannerDevices.Add(deviceInfo);
                                    
                                    // 尝试多种匹配方式
                                    bool idMatch = false;
                                    if (!string.IsNullOrEmpty(scannerId))
                                    {
                                        // 精确匹配
                                        if (id == scannerId)
                                        {
                                            idMatch = true;
                                        }
                                        // 部分匹配（处理不同格式的设备ID）
                                        else if (id.Contains(scannerId) || scannerId.Contains(id))
                                        {
                                            idMatch = true;
                                            Console.WriteLine($"设备ID部分匹配: {id} <-> {scannerId}");
                                        }
                                    }
                                    
                                    if (string.IsNullOrEmpty(scannerId) || idMatch)
                                    {
                                        try
                                        {
                                            var device = deviceInfoType.InvokeMember("Connect", 
                                                BindingFlags.InvokeMethod, null, deviceInfo, null);
                                            Console.WriteLine($"成功连接设备 #{i} (ID匹配)");
                                            return device;
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"连接设备 #{i} 失败: {ex.Message}");
                                            Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"处理设备 #{i} 时出错: {ex.Message}");
                                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"获取设备 #{i} 失败: {ex.Message}");
                        Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                    }
                }

                // 如果没有找到匹配的设备，尝试连接存储的扫描仪设备
                if (scannerDevices.Count > 0)
                {
                    Console.WriteLine($"尝试连接存储的扫描仪设备，数量: {scannerDevices.Count}");
                    foreach (var deviceInfo in scannerDevices)
                    {
                        try
                        {
                            var deviceInfoType = deviceInfo.GetType();
                            var device = deviceInfoType.InvokeMember("Connect", 
                                BindingFlags.InvokeMethod, null, deviceInfo, null);
                            Console.WriteLine("成功连接存储的扫描仪设备");
                            return device;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"连接存储的设备失败: {ex.Message}");
                            Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                            // 继续尝试下一个设备
                        }
                    }
                }

                // 如果仍然没有找到设备，尝试返回第一个可用设备
                if (count > 0)
                {
                    try
                    {
                        var deviceInfo = deviceInfosType.InvokeMember("Item", 
                            BindingFlags.InvokeMethod | BindingFlags.GetProperty, 
                            null, deviceInfos, new object[] { 1 });
                        
                        if (deviceInfo != null)
                        {
                            var deviceInfoType = deviceInfo.GetType();
                            var device = deviceInfoType.InvokeMember("Connect", 
                                BindingFlags.InvokeMethod, null, deviceInfo, null);
                            Console.WriteLine("返回第一个设备作为默认设备");
                            return device;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"获取默认设备失败: {ex.Message}");
                        Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                    }
                }

                Console.WriteLine("未找到匹配的设备");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetDevice异常: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                return null;
            }
        }

        private void SetScanParameters(object device, Type deviceType, ScanRequest request)
        {
            try
            {
                Console.WriteLine("获取设备属性...");
                var properties = deviceType.InvokeMember("Properties", BindingFlags.GetProperty, null, device, null);
                if (properties == null)
                {
                    Console.WriteLine("获取Properties失败: 返回null");
                    return;
                }
                Console.WriteLine("Properties获取成功");
                
                var propertiesType = properties.GetType();
                Console.WriteLine($"Properties类型: {propertiesType.FullName}");
                
                var propertiesCount = (int)propertiesType.InvokeMember("Count", BindingFlags.GetProperty, null, properties, null);
                Console.WriteLine($"属性数量: {propertiesCount}");

                for (int i = 1; i <= propertiesCount; i++)
                {
                    try
                    {
                        var property = propertiesType.InvokeMember("Item", BindingFlags.InvokeMethod, null, properties, new object[] { i });
                        if (property == null)
                        {
                            Console.WriteLine($"属性 #{i}: 返回null");
                            continue;
                        }
                        
                        var propertyType = property.GetType();
                        var propertyId = (int)propertyType.InvokeMember("PropertyID", BindingFlags.GetProperty, null, property, null);
                        
                        Console.WriteLine($"属性 #{i}: ID={propertyId}");

                        switch (propertyId)
                        {
                            case WiaConstants.WIA_IPS_XRES:
                            case WiaConstants.WIA_IPS_YRES:
                                try
                                {
                                    Console.WriteLine($"设置分辨率: {request.Dpi} DPI");
                                    propertyType.InvokeMember("Value", BindingFlags.SetProperty, null, property, new object[] { request.Dpi });
                                    Console.WriteLine("分辨率设置成功");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"设置分辨率失败: {ex.Message}");
                                    // 继续执行，不影响其他参数设置
                                }
                                break;

                            case WiaConstants.WIA_IPA_COLOR_MODE:
                                try
                                {
                                    int colorMode = WiaConstants.WIA_COLOR_MODE_COLOR;
                                    if (request.ColorMode.ToLower() == "grayscale")
                                        colorMode = WiaConstants.WIA_COLOR_MODE_GRAYSCALE;
                                    else if (request.ColorMode.ToLower() == "blackwhite")
                                        colorMode = WiaConstants.WIA_COLOR_MODE_BLACKANDWHITE;
                                    Console.WriteLine($"设置颜色模式: {request.ColorMode} (值: {colorMode})");
                                    propertyType.InvokeMember("Value", BindingFlags.SetProperty, null, property, new object[] { colorMode });
                                    Console.WriteLine("颜色模式设置成功");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"设置颜色模式失败: {ex.Message}");
                                    // 继续执行，不影响其他参数设置
                                }
                                break;

                            case WiaConstants.WIA_DPS_DOCUMENT_HANDLING_SELECT:
                                try
                                {
                                    int paperSource = WiaConstants.WIA_PAPER_SOURCE_FLATBED;
                                    if (request.PaperSource.ToLower() == "feeder")
                                        paperSource = WiaConstants.WIA_PAPER_SOURCE_FEEDER;
                                    Console.WriteLine($"设置纸张来源: {request.PaperSource} (值: {paperSource})");
                                    propertyType.InvokeMember("Value", BindingFlags.SetProperty, null, property, new object[] { paperSource });
                                    Console.WriteLine("纸张来源设置成功");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"设置纸张来源失败: {ex.Message}");
                                    // 继续执行，不影响其他参数设置
                                }
                                break;

                            case WiaConstants.WIA_IPS_BRIGHTNESS:
                                try
                                {
                                    Console.WriteLine($"设置亮度: {request.Brightness}");
                                    propertyType.InvokeMember("Value", BindingFlags.SetProperty, null, property, new object[] { request.Brightness });
                                    Console.WriteLine("亮度设置成功");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"设置亮度失败: {ex.Message}");
                                    // 继续执行，不影响其他参数设置
                                }
                                break;

                            case WiaConstants.WIA_IPS_CONTRAST:
                                try
                                {
                                    Console.WriteLine($"设置对比度: {request.Contrast}");
                                    propertyType.InvokeMember("Value", BindingFlags.SetProperty, null, property, new object[] { request.Contrast });
                                    Console.WriteLine("对比度设置成功");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"设置对比度失败: {ex.Message}");
                                    // 继续执行，不影响其他参数设置
                                }
                                break;

                            case WiaConstants.WIA_IPS_XEXTENT:
                                if (request.Area != null)
                                {
                                    try
                                    {
                                        Console.WriteLine($"设置X范围: {request.Area.Width}");
                                        propertyType.InvokeMember("Value", BindingFlags.SetProperty, null, property, new object[] { request.Area.Width });
                                        Console.WriteLine("X范围设置成功");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"设置X范围失败: {ex.Message}");
                                        // 继续执行，不影响其他参数设置
                                    }
                                }
                                break;

                            case WiaConstants.WIA_IPS_YEXTENT:
                                if (request.Area != null)
                                {
                                    try
                                    {
                                        Console.WriteLine($"设置Y范围: {request.Area.Height}");
                                        propertyType.InvokeMember("Value", BindingFlags.SetProperty, null, property, new object[] { request.Area.Height });
                                        Console.WriteLine("Y范围设置成功");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"设置Y范围失败: {ex.Message}");
                                        // 继续执行，不影响其他参数设置
                                    }
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"处理属性 #{i} 失败: {ex.Message}");
                    }
                }
                Console.WriteLine("扫描参数设置完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置扫描参数失败: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
        }
    }
}