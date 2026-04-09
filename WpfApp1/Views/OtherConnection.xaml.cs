using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    /// <summary>
    /// OtherConnection.xaml 的交互逻辑
    /// </summary>
    public partial class OtherConnection : Window
    {
        public static readonly RoutedCommand ScrollToEndCommand = new RoutedCommand();
        
        public OtherConnection()
        {
            InitializeComponent();
            CommandBindings.Add(new CommandBinding(ScrollToEndCommand, ScrollToEnd_Executed));
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
        private void ScrollToEnd_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Parameter is TextBox textBox)
            {
                textBox.ScrollToEnd();
            }
        }
    }
}