using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WpfApp1.Service.Interfaces;

namespace WpfApp1.Service
{
    /// <summary>
    /// 弹窗服务实现：
    /// 统一处理窗口实例化、Owner挂接与居中显示。
    /// </summary>
    public class DialogService : IDialogService
    {
            /// <summary>
            /// 创建并显示非模态窗口。
            /// </summary>
            public void ShowWindow<TView, TViewModel>(TViewModel viewModel)
                where TView : class, new()
            {
                // 确保在UI线程中执行窗口操作
                if (Application.Current?.Dispatcher != null && !Application.Current.Dispatcher.CheckAccess())
                {
                    Application.Current.Dispatcher.Invoke(() => ShowWindow<TView, TViewModel>(viewModel));
                    return;
                }

                // 1. 实例化窗口
                if (Activator.CreateInstance(typeof(TView)) is not Window window)
                {
                    throw new InvalidOperationException($"类型 {typeof(TView).Name} 必须继承自 System.Windows.Window");
                }

                // 2. 绑定 ViewModel
                window.DataContext = viewModel;

                Window? owner = Application.Current?.MainWindow;
                if (owner != null)
                {
                    window.Owner = owner;
                    window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                }
                else
                {
                    window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                }

                // 使用非模态方式显示窗口，避免阻塞语音识别
                window.Show();
                
                // 确保窗口不会抢占焦点，让语音识别继续工作
                window.Focusable = false;
            }
        }
    }

    




