using System;
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
    public partial class App : Application
    {
        public IServiceProvider? ServiceProvider { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            
            base.OnStartup(e);
            
            try
            {
                NativeLibraryLoader.Initialize();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化原生库失败: {ex.Message}\n{ex.StackTrace}", "启动错误", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Environment.Exit(1);
                return;
            }
            
            var services = new ServiceCollection();
            ConfigureServices(services);
            ServiceProvider = services.BuildServiceProvider();
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show($"UI线程未处理异常: {e.Exception.Message}\n{e.Exception.StackTrace}", 
                "运行时错误", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            MessageBox.Show($"未处理异常: {exception?.Message ?? "未知错误"}\n{exception?.StackTrace}", 
                "致命错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException -= CurrentDomain_UnhandledException;
            DispatcherUnhandledException -= App_DispatcherUnhandledException;
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
