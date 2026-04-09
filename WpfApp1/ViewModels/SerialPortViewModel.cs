using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.ObjectModel;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Windows;
using WpfApp1.Service;

namespace WpfApp1.ViewModels
{
    public partial class SerialPortViewModel : ObservableObject
    {
        private readonly ModbusService _modbusService;
        
        // 连接成功事件
        public event Action? ConnectionSuccess;

        #region 属性定义 
        [ObservableProperty]
        private ObservableCollection<string> _availablePorts = new();

        //自动检测波特率
        [ObservableProperty]
        private Visibility _isAutoDetecting = Visibility.Hidden;
        [ObservableProperty]
        private string _detectionLog = "";
        //端口号
        [ObservableProperty]
        private string? _selectedPort;

        public ObservableCollection<int> AvailableDataBits { get; } = new() { 8, 7, 6 };
        [ObservableProperty]
        private int _dataBit = 8;

        [ObservableProperty]
        private int _timeout = 1000;

        [ObservableProperty]
        private ObservableCollection<int> _availableTimeouts = new() { 1000, 2000, 3000, 4000, 5000 };

        public ObservableCollection<string> AvailableStopBits { get; } = new() { "1", "1.5", "2" };
        [ObservableProperty]
        private string _stopBit = "1";

        public ObservableCollection<int> AvailableBaudRates { get; } = new() { 9600, 19200, 38400, 57600, 115200 };
        [ObservableProperty]
        private int _baudRate = 115200;

        public ObservableCollection<string> AvailableParities { get; } = new() { "无", "奇校验", "偶校验" };
        [ObservableProperty]
        private string _parity = "无";

        [ObservableProperty]
        private string _connectionStatus = "未连接";

        [ObservableProperty]
        private bool _isConnected;

        [ObservableProperty]
        private bool _isAuto=true;
        #endregion

        public SerialPortViewModel(ModbusService modbusService)
        {
            _modbusService = modbusService;
            // 监听ModbusService的IsConnected属性变化
            _modbusService.PropertyChanged += (sender, e) =>
            {
                if (e.PropertyName == nameof(ModbusService.IsConnected))
                {
                    IsConnected = _modbusService.IsConnected;
                    ConnectionStatus = _modbusService.IsConnected ? $"已连接: {SelectedPort} ({BaudRate})" : "已断开";
                    if (!_modbusService.IsConnected) RefreshPorts();
                }
            };
            // 初始化连接状态
            IsConnected = _modbusService.IsConnected;
            ConnectionStatus = _modbusService.IsConnected ? $"已连接: {SelectedPort} ({BaudRate})" : "已断开";
            //刷新端口列表
            RefreshPorts();
        }

        partial void OnSelectedPortChanged(string? value)
        {
            ConnectCommand.NotifyCanExecuteChanged();
        }

        partial void OnIsConnectedChanged(bool value)
        {
            ConnectCommand.NotifyCanExecuteChanged();
            DisconnectCommand.NotifyCanExecuteChanged();
        }

        /// <summary>
        /// 获取可用串口列表
        /// </summary>
        private void RefreshPorts()
        {
            AvailablePorts = _modbusService.RefreshPorts();
            SelectedPort=AvailablePorts.FirstOrDefault();
        }

        /// <summary>
        /// 连接按钮逻辑
        /// </summary>
        [RelayCommand(CanExecute = nameof(CanConnect))]
        private async Task ConnectAsync()
        {
            if (string.IsNullOrEmpty(SelectedPort)) {
                MessageBox.Show("请先选择串口！", "错误");
                return;
            }

            var success = await _modbusService.ConnectAsync(SelectedPort, BaudRate, Parity, DataBit, StopBit, Timeout);
            if (success)
            {
                ConnectionStatus = $"{SelectedPort}连接成功";
                // 触发连接成功事件，通知窗口关闭
                ConnectionSuccess?.Invoke();
            }
            else
            {
                ConnectionStatus = "连接失败";
            }             
            IsAuto = false;
            
        }

        private bool CanConnect() => !IsConnected && !string.IsNullOrEmpty(SelectedPort);

        /// <summary>
        /// 断开按钮逻辑
        /// </summary>
        [RelayCommand(CanExecute = nameof(CanDisconnect))]
        private void Disconnect()
        {
            _modbusService.Disconnect();
            IsAuto=true;
        }

        private bool CanDisconnect() => IsConnected;

        /// <summary>
        /// 批量读取保持寄存器
        /// </summary>
        public async Task<int[]?> ReadMultipleHoldingRegistersAsync(byte slaveAddress, string startAddress, int numberOfPoints)
        {
            return await _modbusService.ReadMultipleHoldingRegistersAsync(slaveAddress, startAddress, numberOfPoints);
        }

        /// <summary>
        /// 写入单个保持寄存器
        /// </summary>
        public async Task<bool> WriteSingleRegisterAsync(byte slaveAddress, string registerAddress, string value)
        {
            return await _modbusService.WriteSingleRegisterAsync(slaveAddress, registerAddress, value);
        }
        /// <summary>
        /// 自动检测波特率
        /// </summary>
        [RelayCommand]
        private async Task AutoDetectBaudRateAsync()
        {
            if (string.IsNullOrEmpty(SelectedPort))
            {
                RefreshPorts();
                MessageBox.Show("请先选择串口");
                return;
            }

            IsAutoDetecting = Visibility.Visible;
            DetectionLog = "正在检测波特率...\n";
            bool found = false;
            int[] baudRates = { 9600, 19200, 38400, 57600, 115200 };

            foreach (int baudRate in baudRates)
            {
                try
                {
                    // 发送一条最小读请求判断该波特率是否可通信
                    var ok = await _modbusService.ProbeDeviceAsync(
                        SelectedPort,
                        baudRate,
                        Parity,
                        DataBit,
                        StopBit,
                        200,
                        1,
                        "1000",
                        1
                    );
                    if (ok)
                    {
                        // 找到可用的波特率
                        BaudRate = baudRate;
                        DetectionLog += $"✓ {baudRate} 波特率连接成功！\n";
                        found = true;
                        break;
                    }
                    DetectionLog += $"✗ {baudRate} 波特率失败\n";
                }
                catch (Exception)
                {
                    // 通信异常按失败记录，继续尝试下一档波特率
                    DetectionLog += $"✗ {baudRate} 波特率失败\n";
                }
            }
            if (found)
            {
                DetectionLog += "\n✓ 已找到设备，波特率已设置为：" + BaudRate + "\n";
            }
            else
            {
                DetectionLog += "\n✗ 暂未寻找到设备\n";
            }
        }

        /// <summary>
        /// 关闭检测界面命令
        /// </summary>
        [RelayCommand]
        private void CloseDetection()
        {
            IsAutoDetecting = Visibility.Hidden;
            DetectionLog = "";
        }

        /// <summary>
        /// 从字符串获取停止位
        /// </summary>
        //private StopBits GetStopBitsFromString(string stopBit)
        //{
        //    return stopBit switch
        //    {
        //        "1.5" => StopBits.OnePointFive,
        //        "2" => StopBits.Two,
        //        _ => StopBits.One
        //    };
        //}

        /// <summary>
        /// 从字符串获取校验位
        /// </summary>
        //private System.IO.Ports.Parity GetParityFromString(string parity)
        //{
        //    return parity switch
        //    {
        //        "奇校验" => System.IO.Ports.Parity.Odd,
        //        "偶校验" => System.IO.Ports.Parity.Even,
        //        _ => System.IO.Ports.Parity.None
        //    };
        //}
    }
}
