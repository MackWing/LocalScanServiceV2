using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.IO;
using Newtonsoft.Json.Serialization;
using LocalScanServiceV2.Services;

namespace LocalScanServiceV2
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // 分析 TwainDotNet 库的 API 结构
            TwainApiInspector.InspectTwainDotNet();
            
            var host = BuildWebHost(args);
            host.Run();
        }

        public static IWebHost BuildWebHost(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
                .UseKestrel()
                .UseContentRoot(Directory.GetCurrentDirectory())
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
                })
                .ConfigureServices(services =>
                {
                    services.AddMvc()
                        .AddJsonOptions(options =>
                        {
                            options.SerializerSettings.ContractResolver = new CamelCasePropertyNamesContractResolver();
                        });

                    services.AddCors(options =>
                    {
                        options.AddDefaultPolicy(policy =>
                            policy.AllowAnyOrigin()
                                .AllowAnyMethod()
                                .AllowAnyHeader());
                    });
                })
                .Configure(app =>
                {
                    app.UseCors();
                    app.UseStaticFiles();
                    app.UseMvc(routes =>
                    {
                        routes.MapRoute(
                            name: "default",
                            template: "{controller=Home}/{action=Index}/{id?}");
                    });
                })
                .UseUrls("http://localhost:9527")
                .Build();
    }
}