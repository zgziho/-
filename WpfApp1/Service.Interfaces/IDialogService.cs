using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp1.Service.Interfaces
{
    /// <summary>
    /// 通用窗口弹出接口：
    /// 统一由服务创建弹窗并绑定对应ViewModel。
    /// </summary>
    public  interface IDialogService
    {
        /// <summary>
        /// 打开指定窗口并绑定指定ViewModel。
        /// </summary>
        void ShowWindow<TView, TViewModel>(TViewModel viewModel)
          where TView : class, new();

        /// <summary>
        /// 打开指定窗口并绑定指定ViewModel，窗口关闭时执行回调。
        /// </summary>
        void ShowWindow<TView, TViewModel>(TViewModel viewModel, Action? onClosed)
          where TView : class, new();
    }
}
