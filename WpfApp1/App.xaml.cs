using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Extensions.DependencyInjection;
using WpfApp1.Service;
using WpfApp1.ViewModels;
using WpfApp1.Service.Interfaces;

namespace WpfApp1
{
    /// <summary>
    /// 应用程序入口：
    /// 负责启动生命周期与全局依赖注入容器初始化。
    /// </summary>
    public partial class App : Application
    {
        public IServiceProvider? ServiceProvider { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            // 统一注册服务和ViewModel，供主窗口和弹窗按需解析
            var services = new ServiceCollection();
            ConfigureServices(services);
            ServiceProvider = services.BuildServiceProvider();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
        }

        private void ConfigureServices(IServiceCollection services)
        {
            // 底层通信与配置服务
            services.AddSingleton<ModbusService>();
            services.AddSingleton<JogConfigService>();
            services.AddSingleton<OtherConnectionService>();
            // 业务ViewModel
            services.AddSingleton<SerialPortViewModel>();
            services.AddSingleton<OtherConnectionViewModel>();
            services.AddSingleton<ExcelViewModel>();
            services.AddSingleton<FirmwareUpdateViewModel>();
            services.AddSingleton<MainWindowViewModel>();
            // 窗口弹出服务
            services.AddSingleton<IDialogService ,DialogService>();
        }


    }
}
