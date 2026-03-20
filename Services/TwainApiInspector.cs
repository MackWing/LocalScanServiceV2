using System;
using System.Reflection;
using System.Linq;

namespace LocalScanServiceV2.Services
{
    public class TwainApiInspector
    {
        public static void InspectTwainDotNet()
        {
            try
            {
                Console.WriteLine("=== 开始分析 TwainDotNet 库 ===");
                
                // 检查 Twain 类型
                var twainType = Type.GetType("TwainDotNet.Twain, TwainDotNet");
                if (twainType != null)
                {
                    Console.WriteLine("\n1. Twain 类:");
                    Console.WriteLine($"   命名空间: {twainType.Namespace}");
                    Console.WriteLine($"   完整名称: {twainType.FullName}");
                    
                    // 检查构造函数
                    var constructors = twainType.GetConstructors();
                    Console.WriteLine("\n   构造函数:");
                    foreach (var ctor in constructors)
                    {
                        var parameters = ctor.GetParameters();
                        var paramStr = string.Join(", ", parameters.Select(p => $"{p.ParameterType.Name} {p.Name}"));
                        Console.WriteLine($"   {twainType.Name}({paramStr})");
                    }
                    
                    // 检查属性
                    var properties = twainType.GetProperties();
                    Console.WriteLine("\n   属性:");
                    foreach (var prop in properties)
                    {
                        Console.WriteLine($"   {prop.PropertyType.Name} {prop.Name}");
                    }
                    
                    // 检查方法
                    var methods = twainType.GetMethods();
                    Console.WriteLine("\n   方法:");
                    foreach (var method in methods)
                    {
                        if (!method.IsSpecialName) // 排除 getter/setter
                        {
                            var parameters = method.GetParameters();
                            var paramStr = string.Join(", ", parameters.Select(p => $"{p.ParameterType.Name} {p.Name}"));
                            Console.WriteLine($"   {method.ReturnType.Name} {method.Name}({paramStr})");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Twain 类型未找到");
                }
                
                // 检查 ScanSettings 类型
                var scanSettingsType = Type.GetType("TwainDotNet.ScanSettings, TwainDotNet");
                if (scanSettingsType != null)
                {
                    Console.WriteLine("\n2. ScanSettings 类:");
                    Console.WriteLine($"   命名空间: {scanSettingsType.Namespace}");
                    Console.WriteLine($"   完整名称: {scanSettingsType.FullName}");
                    
                    // 检查属性
                    var properties = scanSettingsType.GetProperties();
                    Console.WriteLine("\n   属性:");
                    foreach (var prop in properties)
                    {
                        Console.WriteLine($"   {prop.PropertyType.Name} {prop.Name}");
                    }
                }
                else
                {
                    Console.WriteLine("ScanSettings 类型未找到");
                }
                
                // 检查 IWindowsMessageHook 接口
                var messageHookType = Type.GetType("TwainDotNet.IWindowsMessageHook, TwainDotNet");
                if (messageHookType != null)
                {
                    Console.WriteLine("\n3. IWindowsMessageHook 接口:");
                    Console.WriteLine($"   命名空间: {messageHookType.Namespace}");
                    Console.WriteLine($"   完整名称: {messageHookType.FullName}");
                    
                    // 检查方法和属性
                    var members = messageHookType.GetMembers();
                    Console.WriteLine("\n   成员:");
                    foreach (var member in members)
                    {
                        if (member.MemberType == MemberTypes.Method)
                        {
                            var method = (MethodInfo)member;
                            var parameters = method.GetParameters();
                            var paramStr = string.Join(", ", parameters.Select(p => $"{p.ParameterType.Name} {p.Name}"));
                            Console.WriteLine($"   {method.ReturnType.Name} {method.Name}({paramStr})");
                        }
                        else if (member.MemberType == MemberTypes.Property)
                        {
                            var prop = (PropertyInfo)member;
                            Console.WriteLine($"   {prop.PropertyType.Name} {prop.Name}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("IWindowsMessageHook 接口未找到");
                }
                
                Console.WriteLine("\n=== TwainDotNet 库分析完成 ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"分析 TwainDotNet 库时出错: {ex.Message}");
            }
        }
    }
}