using System;
using System.Windows;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    /// <summary>
    /// OtherConnection.xaml 的交互逻辑
    /// </summary>
    public partial class OtherConnection : Window
    {
        public OtherConnection()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}