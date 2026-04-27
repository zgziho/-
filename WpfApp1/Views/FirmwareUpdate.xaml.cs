using System.Windows;

namespace WpfApp1.Views
{
    /// <summary>
    /// FirmwareUpdate.xaml 的交互逻辑
    /// </summary>
    public partial class FirmwareUpdate : Window
    {
        public FirmwareUpdate()
        {
            InitializeComponent();
            this.Closing += FirmwareUpdate_Closing;
        }

        private async void FirmwareUpdate_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (DataContext is ViewModels.FirmwareUpdateViewModel viewModel)
            {
                await viewModel.OnWindowClosing();
            }
        }
    }
}