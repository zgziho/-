using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Globalization;
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
    /// Excel.xaml 的交互逻辑
    /// </summary>
    public partial class Excel : Window
    {

       
        public Excel()
        {
            InitializeComponent();
            var app = (App)Application.Current;
            if (app.ServiceProvider == null)
                throw new InvalidOperationException("ServiceProvider 未初始化。");

           DataContext = app.ServiceProvider.GetRequiredService<ExcelViewModel>();

           Loaded += Excel_Loaded;
           Closing+=Excel_Closing;
        }
        /// <summary>
        /// 打开时自动读取数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
         private void Excel_Loaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is ExcelViewModel viewModel)
            {
                viewModel.StartAutoRead();
            }
        }
        /// <summary>
        /// 关闭时停止自动读取数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Excel_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (DataContext is ExcelViewModel viewModel)
            {
                viewModel.StopAutoRead();
            }
        }

        /// <summary>
        /// 双击修改后触发保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var element = e.Column.GetCellContent(e.Row);
                if (element is FrameworkElement frameworkElement)
                {
                    frameworkElement.BindingGroup?.UpdateSources();
                }

                var vm = DataContext as ExcelViewModel;
                if (vm != null)
                {
                    if (e.Column.Header?.ToString() == "数据值")
                    {
                        // 获取当前编辑的行数据
                        var selectedRow = e.Row.Item as WpfApp1.Models.Pars;
                        if (selectedRow != null)
                        {
                            // 尝试解析寄存器地址和数据值
                            if (TryToHexWord(selectedRow.ParsID, true, out var registerAddressHex) &&
                                TryToHexWord(selectedRow.ParsVA, false, out var valueHex) && selectedRow.ParsSX == "读写")
                            {
                                try
                                {
                                    // 根据当前连接选择协议：串口(Modbus) 或 CAN
                                    if (vm._otherConnectionService != null && vm._otherConnectionService.IsStarted)
                                    {
                                        // CAN 写：将 valueHex (4 字符，如 "00AF") 转为 2 字节并构建负载
                                        var startAddress = int.Parse(registerAddressHex, System.Globalization.NumberStyles.HexNumber);
                                        var hi = Convert.ToByte(valueHex.Substring(0, 2), 16);
                                        var lo = Convert.ToByte(valueHex.Substring(2, 2), 16);
                                        var data = new byte[] { hi, lo };
                                        var payload = vm._otherConnectionService.BuildWritePayload(startAddress, data);
                                        // 发送可能阻塞，在线程池中执行
                                        var result = await System.Threading.Tasks.Task.Run(() => vm._otherConnectionService.SendWriteMessage(payload));
                                        if (!result.Succeeded)
                                        {
                                            throw new System.Exception(result.Message ?? "CAN 发送失败");
                                        }
                                    }
                                    else
                                    {
                                        // 串口(Modbus) 写：发送单个寄存器
                                        await vm._modbusService.WriteSingleRegisterAsync(1, registerAddressHex, valueHex);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Windows.MessageBox.Show($"发送数据到设备失败：{ex.Message}", "发送失败",
                                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                                }
                            }
                            else
                            {
                                System.Windows.MessageBox.Show("数据格式错误，无法发送到设备", "格式错误",
                                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                            }
                        }
                    }
                    else
                    {
                        // 其他列，保存到文件
                        vm.SaveData();
                    }
                }
            }
        }

        /// <summary>
        /// 导航栏菜单切换
        /// </summary>
        private void NavListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataContext is ExcelViewModel viewModel && NavListBox.SelectedItem is ListBoxItem selectedItem)
            {
                if (int.TryParse(selectedItem.Tag?.ToString(), out int index))
                {
                    viewModel.SelectMenuCommand.Execute(index);
                }
            }
        }
        /// <summary>
        /// 将数据转换为合适的
        /// </summary>
        /// <param name="text"></param>
        /// <param name="hexOnly"></param>
        /// <param name="hex"></param>
        /// <returns></returns>
        private static bool TryToHexWord(string? text, bool hexOnly, out string hex)
        {
            hex = string.Empty;
            text = (text ?? string.Empty).Trim();

            if (hexOnly)
            {
                if (!int.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var address))
                    return false;
                if (address < 0 || address > 0xFFFF)
                    return false;
                hex = address.ToString("X4");
                return true;
            }

            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
                return false;
            hex = (value & 0xFFFF).ToString("X4");
            return true;
        }
    }
}

