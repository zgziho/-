using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    /// <summary>
    /// Serialport.xaml 的交互逻辑
    /// </summary>
    public partial class Serialport : Window
    {
        public Serialport()
        {
            InitializeComponent();
            
            // 在窗口加载完成后订阅事件，确保DataContext已设置
            this.Loaded += OnWindowLoaded;
        }
        
        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            // 获取ViewModel并订阅连接成功事件
            if (DataContext is SerialPortViewModel viewModel)
            {
                viewModel.ConnectionSuccess += OnConnectionSuccess;
            }
        }
        
        private void OnConnectionSuccess()
        {
            // 连接成功后关闭窗口
            this.Close();
        }
        
        protected override void OnClosed(EventArgs e)
        {
            // 清理事件订阅
            if (DataContext is SerialPortViewModel viewModel)
            {
                viewModel.ConnectionSuccess -= OnConnectionSuccess;
            }
            base.OnClosed(e);
        }
    }
}


