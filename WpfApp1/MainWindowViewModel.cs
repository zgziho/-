using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using OpenTK.Graphics.GL;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using WpfApp1.Models;
using WpfApp1.Service;
using WpfApp1.Service.Interfaces;
using WpfApp1.ViewModels;
using WpfApp1.Views;

namespace WpfApp1
{
    public partial class MainWindowViewModel : ObservableObject, IDisposable
    {
        /// <summary>
        /// 实时读取模式枚举
        /// </summary>
        private enum RealtimeReadMode
        {
            /// <summary>无读取模式</summary>
            None,
            /// <summary>串口读取模式</summary>
            Serial,
            /// <summary>CAN读取模式</summary>
            Can
        }

        /// <summary>
        /// 写入协议模式枚举
        /// </summary>
        private enum WriteProtocolMode
        {
            /// <summary>无写入模式</summary>
            None,
            /// <summary>串口写入模式</summary>
            Serial,
            /// <summary>CAN写入模式</summary>
            Can
        }

        /// <summary>
        /// 服务提供者，用于获取依赖注入的服务
        /// </summary>
        public readonly IServiceProvider _serviceProvider;

        /// <summary>
        /// 对话框服务，用于显示窗口
        /// </summary>
        private readonly IDialogService _dialogService;
        /// <summary>
        /// Excel视图模型，用于处理Excel相关操作
        /// </summary>
        private readonly ExcelViewModel _excelViewModel;
        /// <summary>
        /// Modbus服务，用于串口通信
        /// </summary>
        private readonly ModbusService _modbusService;
        /// <summary>
        /// 其他连接服务，用于CAN通信
        /// </summary>
        private readonly OtherConnectionService _otherConnectionService;
        /// <summary>
        /// 点动/使能配置服务：从 JogConfig.xlsx 加载名称与地址映射
        /// </summary>
        private readonly JogConfigService _jogConfigService;
        /// <summary>
        /// 运行期配置缓存：Name -> 配置项，便于快速按名称查地址和值
        /// </summary>
        private readonly Dictionary<string, JogConfigItem> _jogConfigs = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// 当前控制模式对应的“给定”名称
        /// </summary>
        private string _currentSetpointConfigName = "电流给定";
        /// <summary>
        /// 长按点动过程标记：用于避免重复下发“开指令”
        /// </summary>
        private bool _isJogPressing;

        /// <summary>
        /// 串口视图模型
        /// </summary>
        public SerialPortViewModel SerialPortViewModel => _serviceProvider.GetRequiredService<SerialPortViewModel>();

        /// <summary>
        /// 是否打开了参数表
        /// </summary>
        bool IsOpenExcel = false;
        /// <summary>
        /// 定时器，用于定期更新UI
        /// </summary>
        private DispatcherTimer timer;

        /// <summary>
        /// modbus读取专用线程
        /// </summary>
        private Thread _modbusReadThread;
        /// <summary>
        /// 紧急停止令牌源
        /// </summary>
        private CancellationTokenSource _modbusReadCts;
        /// <summary>
        /// 读取间隔事件
        /// </summary>
        private ManualResetEventSlim _modbusReadEvent;
        /// <summary>
        /// 是否在读取数据
        /// </summary>
        private volatile bool _isModbusReading = false;
        /// <summary>
        /// 读取间隔时间（毫秒）
        /// </summary>
        private volatile int _readIntervalMs = 100;

        /// <summary>
        /// 预加载的读取参数列表
        /// </summary>
        private List<Pars> _preloadedReadParameters = new List<Pars>();
        /// <summary>
        /// 地址到参数的映射表
        /// </summary>
        private Dictionary<int, Pars> _addressToParsMap = new Dictionary<int, Pars>();
        /// <summary>
        /// 上次寄存器值缓存
        /// </summary>
        private Dictionary<int, int> _lastRegisterValues = new Dictionary<int, int>();

        /// <summary>
        /// 控制模式选择是否已成功更新过一次
        /// </summary>
        private bool _controlModeSelectionUpdated = false;


        /// <summary>
        /// 选中的控制模式
        /// </summary>
        [ObservableProperty]
        private string _selectedOption = "--未选择--";

        /// <summary>
        /// 选中的分组视图模型
        /// </summary>
        [ObservableProperty]
        private object _selectedGroup;
        /// <summary>
        /// 各模式的给定值
        /// </summary>
        [ObservableProperty]
        private string _currentSetpointValue = "0";

        /// <summary>
        /// 各模式的给定值存储
        /// </summary>
        private Dictionary<string, string> _setpointValues = new()
        {
            { "电流给定", "0" },
            { "速度给定", "0" },
            { "位置给定", "0" },
            { "占空比给定", "0" },
            { "旋转频率", "0" },
        };

        /// <summary>
        /// 所有参数的存储字典（按参数名存储）
        /// </summary>
        private Dictionary<string, string> _parameterValues = new()
        {
            // 电流环参数
            { "Dkp", "" },
            { "Qkp", "" },
            { "CurKi", "" },  // 电流环Ki
            
            // 速度环参数
            { "SpeedKp", "" },  // 速度环Kp
            { "SpeedKi", "" },  // 速度环Ki
            { "Tff", "" },
            { "SpeedBand", "" },
            
            // 位置环参数
            { "PosKp", "" },  // 位置环Kp
            { "Vff", "" },
            { "CurrentPosition", "" },
            
            // VF模式参数
            // "占空比给定"和"旋转频率"已在_setpointValues中存储
            { "DutyCycle", "" },
            
            // 变频器参数
            { "旋转频率给定", "" },
            { "电流比率给定", "" },
            { "电流数字给定", "" },
            { "电流峰值给定", "" }
        };

        /// <summary>
        /// 通道1选择
        /// </summary>
        [ObservableProperty]
        private string _channel1Selected = "无";

        /// <summary>
        /// 通道1选择变更时的处理
        /// </summary>
        partial void OnChannel1SelectedChanged(string value)
        {
            UpdateChannelSelectionStatus();
        }

        /// <summary>
        /// 通道2选择
        /// </summary>
        [ObservableProperty]
        private string _channel2Selected = "无";

        /// <summary>
        /// 通道2选择变更时的处理
        /// </summary>
        partial void OnChannel2SelectedChanged(string value)
        {
            UpdateChannelSelectionStatus();
        }

        /// <summary>
        /// 通道3选择
        /// </summary>
        [ObservableProperty]
        private string _channel3Selected = "无";

        /// <summary>
        /// 通道3选择变更时的处理
        /// </summary>
        partial void OnChannel3SelectedChanged(string value)
        {
            UpdateChannelSelectionStatus();
        }

        /// <summary>
        /// 通道4选择
        /// </summary>
        [ObservableProperty]
        private string _channel4Selected = "无";

        /// <summary>
        /// 通道4选择变更时的处理
        /// </summary>
        partial void OnChannel4SelectedChanged(string value)
        {
            UpdateChannelSelectionStatus();
        }

        /// <summary>
        /// 示波器数据管理 - 高性能环形缓冲区实现
        /// 每个通道独立存储2000个数据点，支持高速数据接收（2ms/组）
        /// </summary>
        private class CircularBuffer
        {
            private readonly double[] _buffer;
            private int _writeIndex = 0;
            private int _count = 0;
            private readonly object _lock = new object();
            public int Capacity { get; }

            public CircularBuffer(int capacity)
            {
                _buffer = new double[capacity];
                Capacity = capacity;
            }

          

            public void Write(double[] items)
            {
                if (items == null || items.Length == 0) return;
                
                lock (_lock)
                {
                    int length = items.Length;
                    int capacity = Capacity;

                    if (length <= capacity - _writeIndex)
                    {
                        Array.Copy(items, 0, _buffer, _writeIndex, length);
                       
                    }
                    else
                    {
                        var count = capacity - _writeIndex;
                        Array.Copy(items, 0, _buffer, _writeIndex, count);
                        Array.Copy(items, count, _buffer, 0, length - count);
                    }
                    
                    _writeIndex = (_writeIndex + length) % capacity;
                    _count = Math.Min(_count + length, capacity);
                }
            }
            /// <summary>
            /// 从指定读取索引开始读取指定数量的数据点
            /// </summary>
            /// <param name="readIndex">读取索引</param>
            /// <param name="count">要读取的数据点数量</param>
            /// <returns>读取的数据点数组</returns>
            public double[] ReadFromIndex(ref int readIndex, int count)
            {
                lock (_lock)
                {
                    if (count> _count) return Array.Empty<double>();

                    var result = new double[count];
                    
                    for (int i = 0; i < count; i++)
                    {
                        result[i] = _buffer[readIndex];
                        readIndex = (readIndex + 1) % Capacity;
                    }
                    _count -= count;
                    return result;
                }
            }

            public void Clear()
            {
                lock (_lock)
                {
                    _writeIndex = 0;
                    _count = 0;
                }
            }
        }



        /// <summary>
        /// 示波器数据缓存 - 每个通道独立的高性能环形缓冲区
        /// </summary>
        private readonly Dictionary<int, CircularBuffer> _oscilloscopeBuffers = new()
        {
            { 1, new CircularBuffer(50000) },  // 通道1：2000个点容量
            { 2, new CircularBuffer(50000) },  // 通道2：2000个点容量
            { 3, new CircularBuffer(50000) },  // 通道3：2000个点容量
            { 4, new CircularBuffer(50000) }   // 通道4：2000个点容量
        };

        /// <summary>
        /// 示波器数据锁 - 确保多线程安全访问
        /// </summary>
        private readonly object[] _oscilloscopeLocks = new object[] { new(), new(), new(), new(), new() };

        /// <summary>
        /// 示波器读取索引 - 每个通道独立的读取索引，确保数据连续性
        /// </summary>
        private readonly Dictionary<int, int> _readIndices = new()
        {
            { 1, 0 },  // 通道1初始读取索引
            { 2, 0 },  // 通道2初始读取索引
            { 3, 0 },  // 通道3初始读取索引
            { 4, 0 }   // 通道4初始读取索引
        };

        /// <summary>
        /// 下拉框状态：是否有通道被选择
        /// </summary>
        private bool _hasChannelSelected = false;
        /// <summary>
        /// 当前模式对应给定的显示名称
        /// </summary>
        [ObservableProperty]
        private string _setpointDisplayName = "电流给定";

        /// <summary>
        /// 检查并更新下拉框选择状态，控制示波器开启/关闭
        /// </summary>
        private async void UpdateChannelSelectionStatus()
        {
            bool hasChannelSelected = Channel1Selected != "无"
                || Channel2Selected != "无"
                || Channel3Selected != "无"
                || Channel4Selected != "无";

            // 更新通道选择状态
            _hasChannelSelected = hasChannelSelected;

            // 处理通道选择索引写入逻辑
            await HandleChannelSelectionIndex();

            // 处理示波器开启/关闭逻辑
            await HandleOscilloscopeControl(hasChannelSelected);



        }

        /// <summary>
        /// 使能按钮动态文本：使能 / 关使能
        /// </summary>
        [ObservableProperty]
        private string _enableButtonText = "使能";
        /// <summary>
        /// 正向按钮动态文本：正向点动 / 正向点动中 / 反向单点
        /// </summary>
        [ObservableProperty]
        private string _forwardJogButtonText = "正向点动";
        /// <summary>
        /// 反向按钮动态文本：反向点动 / 反向点动中 / 正向单点
        /// </summary>
        [ObservableProperty]
        private string _reverseJogButtonText = "反向点动";
        /// <summary>
        /// 使能锁存状态：true 时进入“单点模式”，false 时是“长按点动模式”
        /// </summary>
        [ObservableProperty]
        private bool _isEnableLatched;
        /// <summary>
        /// 模式选项数据源
        /// </summary>
        public List<string> GroupOptions { get; } = new() { "--未选择--", "电流环", "速度环", "位置环", "VF模式", "变频器", "自学习", "过流测试", "转矩环频带" };

        /// <summary>
        /// 通道选项数据源
        /// </summary>
        public List<string> ChannelOptions { get; } = new() { "无", "母线电压", "电流给定", "电流反馈", "速度给定", "速度反馈", "位置给定",
            "位置反馈", "U相采样", "V相采样", "Alpha","Beta","D轴电流","Q轴电流","Vd","Vq","Ualpha","UBeta","电流电角度",
        "反馈电角度","系统温度","电机温度"};

        /// <summary>
        /// 选中模式变更时的处理
        /// </summary>
        partial void OnSelectedOptionChanged(string value)
        {
            // 保存当前模式的给定值和所有参数
            if (!string.IsNullOrEmpty(_currentSetpointConfigName))
            {
                _setpointValues[_currentSetpointConfigName] = CurrentSetpointValue;
            }
            
            // 保存当前模式的所有参数值
            SaveCurrentGroupParameters();

            // 处理"--未选择--"选项
            if (value == "--未选择--")
            {
                SelectedGroup = null;
                _currentSetpointConfigName = string.Empty;
                SetpointDisplayName = string.Empty;
                // 发送控制模式索引
                _ = HandleControlModeSelection();
                return;
            }

            switch (value)
            {
                case "电流环":
                    SelectedGroup = new ViewModels.GroupAViewModel();
                    // 电流环模式下，点动使用“电流给定”作为写入目标
                    _currentSetpointConfigName = "电流给定";
                    break;
                case "速度环":
                    SelectedGroup = new ViewModels.GroupBViewModel();
                    // 速度环模式下，点动使用“速度给定”作为写入目标
                    _currentSetpointConfigName = "速度给定";
                    break;
                case "位置环":
                    SelectedGroup = new ViewModels.GroupCViewModel();
                    // 位置环模式下，点动使用“位置给定”作为写入目标
                    _currentSetpointConfigName = "位置给定";
                    break;
                case "VF模式":
                    SelectedGroup = new ViewModels.GroupDViewModel();
                    // VF模式下，点动使用“占空比给定”作为写入目标
                    _currentSetpointConfigName = "占空比给定";
                    break;
                case "变频器":
                    SelectedGroup = new ViewModels.GroupEViewModel();
                    break;
                case "自学习":
                    SelectedGroup = new ViewModels.GroupFViewModel();
                    _currentSetpointConfigName = string.Empty;
                    SetpointDisplayName = string.Empty;
                    break;
                case "过流测试":
                    SelectedGroup = new ViewModels.GroupGViewModel();
                    _currentSetpointConfigName = "占空比给定";
                    break;
                case "转矩环频带":
                    SelectedGroup = new ViewModels.GroupHViewModel();
                    _currentSetpointConfigName = "电流给定";
                    break;
            }

            // 加载新模式的给定值
            if (!string.IsNullOrEmpty(_currentSetpointConfigName) && _setpointValues.ContainsKey(_currentSetpointConfigName))
            {
                CurrentSetpointValue = _setpointValues[_currentSetpointConfigName];
            }
            
            // 加载新模式的参数值
            LoadCurrentGroupParameters();

            SetpointDisplayName = _currentSetpointConfigName;

            // 发送控制模式索引
            _ = HandleControlModeSelection();
        }

        /// <summary>
        /// 保存当前模式的所有参数值
        /// </summary>
        private void SaveCurrentGroupParameters()
        {
            if (SelectedGroup == null) return;

            switch (SelectedGroup)
            {
                case GroupAViewModel groupA:
                    _parameterValues["Qkp"] = groupA.Qkp;
                    _parameterValues["CurKi"] = groupA.Ki;
                    break;
                case GroupBViewModel groupB:
                    _parameterValues["SpeedKp"] = groupB.Kp;
                    _parameterValues["SpeedKi"] = groupB.Ki;
                    _parameterValues["Tff"] = groupB.Tff;
                    _parameterValues["SpeedBand"] = groupB.SpeedBand;
                    break;
                case GroupCViewModel groupC:
                    _parameterValues["PosKp"] = groupC.Kp;
                    _parameterValues["Vff"] = groupC.Vff;
                    _parameterValues["CurrentPosition"] = groupC.CurrentPosition;
                    break;
                case GroupDViewModel groupD:
                    _parameterValues["DutyCycle"] = groupD.DutyCycle;
                    break;
                case GroupEViewModel groupE:
                    _parameterValues["旋转频率给定"] = groupE.RotateFreq;
                    _parameterValues["电流比率给定"] = groupE.CurrentRatio;
                    _parameterValues["电流数字给定"] = groupE.CurrentDigitalRef;
                    _parameterValues["电流峰值给定"] = groupE.CurrentPeak;
                    break;
                case GroupFViewModel groupF:
                    break;
                case GroupGViewModel groupG:
                    _parameterValues["占空比给定"] = groupG.DutyCycle;
                    _parameterValues["旋转频率"] = groupG.RotateFreq;
                    break;
                case GroupHViewModel groupH:
                    _parameterValues["电流给定"] = groupH.CurrentRef;
                    _parameterValues["旋转频率给定"] = groupH.SignalFreq;
                    break;
            }
        }

        /// <summary>
        /// 加载当前模式的所有参数值
        /// </summary>
        private void LoadCurrentGroupParameters()
        {
            if (SelectedGroup == null) return;

            switch (SelectedGroup)
            {
                case GroupAViewModel groupA:
                    groupA.Dkp = _parameterValues["Dkp"] ?? "";
                    groupA.Qkp = _parameterValues["Qkp"] ?? "";
                    groupA.Ki = _parameterValues["CurKi"] ?? "";
                    break;
                case GroupBViewModel groupB:
                    groupB.Kp = _parameterValues["SpeedKp"] ?? "";
                    groupB.Ki = _parameterValues["SpeedKi"] ?? "";
                    groupB.Tff = _parameterValues["Tff"] ?? "";
                    groupB.SpeedBand = _parameterValues["SpeedBand"] ?? "";
                    break;
                case GroupCViewModel groupC:
                    groupC.Kp = _parameterValues["PosKp"] ?? "";
                    groupC.Vff = _parameterValues["Vff"] ?? "";
                    groupC.CurrentPosition = _parameterValues["CurrentPosition"] ?? "";
                    break;
                case GroupDViewModel groupD:
                    groupD.DutyCycle = _parameterValues["DutyCycle"] ?? "";
                    break;
                case GroupEViewModel groupE:
                    groupE.RotateFreq = _parameterValues["旋转频率给定"] ?? "";
                    groupE.CurrentRatio = _parameterValues["电流比率给定"] ?? "";
                    groupE.CurrentDigitalRef = _parameterValues["电流数字给定"] ?? "";
                    groupE.CurrentPeak = _parameterValues["电流峰值给定"] ?? "";
                    break;
                case GroupFViewModel groupF:
                    break;
                case GroupGViewModel groupG:
                    groupG.DutyCycle = _parameterValues["占空比给定"] ?? "";
                    groupG.RotateFreq = _parameterValues["旋转频率"] ?? "";
                    break;
                case GroupHViewModel groupH:
                    groupH.CurrentRef = _parameterValues["电流给定"] ?? "";
                    groupH.SignalFreq = _parameterValues["旋转频率给定"] ?? "";
                    break;
            }
        }

        /// <summary>
        /// 提交指定参数到设备
        /// </summary>
        /// <param name="parameterName">参数名称</param>
        /// <param name="parameterValue">参数值</param>
        /// <returns>操作是否成功</returns>
        public async Task<bool> CommitParameterAsync(string parameterName, string parameterValue)
        {
            if (!TryGetConfig(parameterName, out var config))
                return false;

            return await WriteConfigValueAsync(config, parameterValue);
        }



        #region 文本框参数
        /// <summary>
        /// 母线电压
        /// </summary>
        [ObservableProperty]
        private string _voltage = "/";
        /// <summary>
        /// 驱动器温度
        /// </summary>
        [ObservableProperty]
        private string _driveTemperature = "/";
        /// <summary>
        /// 电机温度
        /// </summary>
        [ObservableProperty]
        private string _motorTemperature = "/";
        /// <summary>
        /// 电流给定
        /// </summary>
        [ObservableProperty]
        private string _currentSet = "/";
        /// <summary>
        /// 电流反馈峰值
        /// </summary>
        [ObservableProperty]
        private string _currentFeedbackPeak = "/";
        /// <summary>
        /// 电流有效值
        /// </summary>
        [ObservableProperty]
        private string _currentRMS = "/";
        /// <summary>
        /// 速度给定
        /// </summary>
        [ObservableProperty]
        private string _speedSet = "/";
        /// <summary>
        /// 速度反馈
        /// </summary>
        [ObservableProperty]
        private string _speedFeedback = "/";
        /// <summary>
        /// 位置给定
        /// </summary>
        [ObservableProperty]
        private string _positionSet = "/";
        /// <summary>
        /// 位置反馈
        /// </summary>
        [ObservableProperty]
        private string _positionFeedback = "/";
        #endregion

        private bool _isFirstRead = true;
        /// <summary>
        /// 主窗口视图模型构造函数
        /// </summary>
        /// <param name="serviceProvider">服务提供者</param>
        /// <param name="dialogService">对话框服务</param>
        /// <param name="excelViewModel">Excel视图模型</param>
        /// <param name="modbusService">Modbus服务</param>
        /// <param name="otherConnectionService">其他连接服务</param>
        public MainWindowViewModel(IServiceProvider serviceProvider, IDialogService dialogService, ExcelViewModel excelViewModel, ModbusService modbusService, OtherConnectionService otherConnectionService)
        {
            _serviceProvider = serviceProvider;
            _dialogService = dialogService;
            _excelViewModel = excelViewModel;
            _modbusService = modbusService;
            _otherConnectionService = otherConnectionService;
            _jogConfigService = _serviceProvider.GetRequiredService<JogConfigService>();

            // 预加载读取参数配置
            PreloadReadParameters();
            //读取配置表数据
            LoadJogConfigs();

            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(500);
            timer.Tick += Timer_Tick;
            timer.Start();

            // 初始化专用Modbus读取线程
            InitializeModbusReadThread();

            Debug.WriteLine($"[MainWindowViewModel] 专用Modbus读取线程已启动，预加载 {_preloadedReadParameters.Count} 个参数");
        }

        /// <summary>
        /// 预加载读取参数配置（筛选数据）
        /// </summary>
        private void PreloadReadParameters()
        {
            try
            {
                // 从Excel配置中预加载"读"项参数
                _preloadedReadParameters = _excelViewModel._allPers
                    .Where(p => !string.IsNullOrEmpty(p.ParsID) && p.ParsSX?.Trim() == "读")
                    .OrderBy(p => ParseAddress(p.ParsID))
                    .ToList();

                // 创建地址到参数的映射表
                _addressToParsMap.Clear();
                foreach (var param in _preloadedReadParameters)
                {
                    int address = ParseAddress(param.ParsID);
                    if (address > 0)
                    {
                        _addressToParsMap[address] = param;
                    }
                }

                Debug.WriteLine($"[Preload] 预加载完成: {_preloadedReadParameters.Count} 个参数，{_addressToParsMap.Count} 个地址映射");
                // 记录预加载参数的最小起始地址，供 CAN 实时数据写入缓存使用
                if (_preloadedReadParameters.Count > 0)
                {
                    _preloadedMinAddress = ParseAddress(_preloadedReadParameters.First().ParsID);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Preload] 预加载失败: {ex.Message}");
            }

        }

        /// <summary>
        /// 预加载所有参数配置（未筛选数据）
        /// </summary>
        private void PreloadAllParameters()
        {
            try
            {
                // 从Excel配置中预加载所有参数，不进行筛选
                var allParameters = _excelViewModel._allPers
                    .Where(p => !string.IsNullOrEmpty(p.ParsID))
                    .OrderBy(p => ParseAddress(p.ParsID))
                    .ToList();

                Debug.WriteLine($"[PreloadAll] 预加载所有参数完成: {allParameters.Count} 个参数");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[PreloadAll] 预加载所有参数失败: {ex.Message}");
            }
        }

        private int _preloadedMinAddress = 0;
        private bool _isPassiveCanListening = false;

        private void StartPassiveCanListening()
        {
            if (_isPassiveCanListening)
                return;
            _isPassiveCanListening = true;
            try
            {
                _otherConnectionService.RealtimeFrameReceived += OnCanRealtimeFrame;
                _otherConnectionService.StartPassiveReceive();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CanListen] 启动被动接收失败: {ex.Message}");
            }
        }

        private void StopPassiveCanListening()
        {
            if (!_isPassiveCanListening)
                return;
            _isPassiveCanListening = false;
            try
            {
                _otherConnectionService.RealtimeFrameReceived -= OnCanRealtimeFrame;
                _otherConnectionService.StopPassiveReceive();
            }
            catch { }
        }

        /// <summary>
        /// 将接收到的数据进行处理
        /// </summary>
        /// <param name="payload"></param>
        private void OnCanRealtimeFrame(byte[] payload)
        {
            try
            {
                if (payload == null || payload.Length == 0)
                    return;

                // 首字节为 0x31..0x34，代表通道 1..4；也可能为 0x30 表示其他类型
                byte first = payload[0];
                if (first < 0x31 || first > 0x34)
                {
                    if (first == 0x30)
                    {
                        int realtimeLen = payload[1];

                        int idx = 2;
                        if (realtimeLen < 0) realtimeLen = 0;

                        if (realtimeLen > 0 && payload.Length >= idx + realtimeLen)
                        {
                            var realtimeBytes = new byte[realtimeLen];
                            Array.Copy(payload, idx, realtimeBytes, 0, realtimeLen);

                            // 将 realtimeBytes 解析为寄存器数组（每两个字节为一寄存器，高字节在前）
                            int regCount = realtimeBytes.Length / 2;
                            if (regCount > 0)
                            {
                                var regs = new int[regCount];
                                for (int i = 0; i < regCount; i++)
                                {
                                    int lo = realtimeBytes[i * 2];
                                    int hi = realtimeBytes[i * 2 + 1];
                                    regs[i] = (hi << 8) | lo;
                                }
                                // 写入到 Modbus 缓存（使用预加载的最小地址）
                                _modbusService?.UpdateCacheFromExternal(_preloadedMinAddress, regs);
                            }
                        }
                        // 非通道实时包，忽略
                        return;
                    }
                }



                int channel = first - 0x30; // 1..4
                
                if (channel<=0||channel>4) 
                    return;

                // 处理示波器数据
                ProcessOscilloscopeData(payload, channel);
            }
            
            catch (Exception ex)
            {
                Debug.WriteLine($"[CanFrameHandler] Error: {ex.Message}");
            }
            }

        /// <summary>
        /// 初始化专用Modbus读取线程
        /// </summary>
        private void InitializeModbusReadThread()
        {
            _modbusReadCts = new CancellationTokenSource();
            _modbusReadEvent = new ManualResetEventSlim(false);
            _modbusReadThread = new Thread(ModbusReadThreadWorker)
            {//线程名字
                Name = "ModbusReadThread",
                //程序关闭后自动关闭
                IsBackground = true,
                //高优先级
                Priority = ThreadPriority.Highest
            };
            _modbusReadThread.Start();
            
            // 启动后立即开始读取
            StartHighFrequencyRead();
        }


       
        /// <summary>
        /// 专用Modbus读取线程的工作方法
        /// </summary>
        private async void ModbusReadThreadWorker()
        {
            var token = _modbusReadCts.Token;

            while (!token.IsCancellationRequested)
            {
                try
                {
                    _modbusReadEvent.Wait(_readIntervalMs, token);
                    var readMode = GetRealtimeReadMode();

                    if (_isModbusReading && readMode != RealtimeReadMode.None)
                    {
                        if (_isFirstRead)
                        {
                            // 第一次读取时，读取所有数据到缓存
                            await ExecuteHighFrequencyRead(readMode, true);
                            // 重新预加载参数（确保使用最新数据）
                            PreloadReadParameters();
                            // 标记第一次读取已完成
                            _isFirstRead = false;
                        }
                        else
                        {
                            // 后续读取使用筛选的数据
                            await ExecuteHighFrequencyRead(readMode, false);
                        }
                    }
                    else if(readMode == RealtimeReadMode.None)
                    {

                    }
                    else
                    {
                        //Debug.WriteLine($"[ModbusReadThread] 跳过读取 - IsReading: {_isModbusReading}, Mode: {readMode}");
                    }

                    _modbusReadEvent.Reset();
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ModbusReadThread] Error: {ex.Message}");
                    await Task.Delay(1); // 短暂休眠避免死循环
                }
            }

        }

        /// <summary>
        /// 获取实时读取模式
        /// </summary>
        /// <returns>当前的实时读取模式</returns>
        private RealtimeReadMode GetRealtimeReadMode()
        {
            // 如果串口已连接，使用串口读取模式
            if (_modbusService.IsConnected)
            {
                return RealtimeReadMode.Serial;
            }

            // 如果CAN已启动，使用CAN读取模式
            if (_otherConnectionService.IsStarted)
            {
                return RealtimeReadMode.Can;
            }
            
            // 无可用连接
            return RealtimeReadMode.None;
        }

        /// <summary>
        /// 获取写入协议模式
        /// </summary>
        /// <returns>当前的写入协议模式</returns>
        private WriteProtocolMode GetWriteProtocolMode()
        {
            // 如果串口已连接，使用串口写入模式
            if (_modbusService.IsConnected)
            {
                return WriteProtocolMode.Serial;
            }

            // 如果CAN已启动，使用CAN写入模式
            if (_otherConnectionService.IsStarted)
            {
                return WriteProtocolMode.Can;
            }

            // 无可用连接
            return WriteProtocolMode.None;
        }



      
        /// <summary>
        /// 从缓存中获取最大电流值
        /// </summary>
        /// <returns>最大电流值</returns>
        private double GetMaxCurrentFromCache()
        {
            // 这里需要根据实际情况从缓存中获取最大电流值
            // 假设最大电流值存储在某个特定地址
            // 示例：从地址0x1000获取最大电流值
            string maxCurrentStr = _modbusService?.GetCachedValue(0x1000) ?? "100";
            if (double.TryParse(maxCurrentStr, out double maxCurrent))
            {
                return maxCurrent;
            }
            return 100; // 默认值
        }

        /// <summary>
        /// 从缓存更新UI界面（在定时器中调用）
        /// </summary>
        private void UpdateUIFromCache()
        {
            if (_preloadedReadParameters.Count == 0)
                return;

            var updateList = new List<(Pars parsItem, string parameterName, string value)>();
            
            // 收集需要更新的参数（从缓存中读取）
            foreach (var param in _preloadedReadParameters)
            {
                int address = ParseAddress(param.ParsID);
                if (address > 0)
                {
                    // 从modbusService缓存中获取最新值，而不是使用预加载的默认值
                    string cachedValue = _modbusService?.GetCachedValue(address) ?? param.ParsVA;
                    if (!string.IsNullOrEmpty(cachedValue))
                    {
                        updateList.Add((param, param.ParsNM, cachedValue));
                    }
                }
            }
            
            // 批量UI更新
            Application.Current?.Dispatcher.Invoke(() =>
            {
                // 更新常规参数
                foreach (var (parsItem, parameterName, value) in updateList)
                {
                    if (!string.IsNullOrEmpty(parameterName))
                    {
                        UpdateViewModelProperty(parameterName, value);
                    }
                }
                
                // 更新控制模式下拉框选择
                UpdateControlModeSelection();
            });
        }

        /// <summary>
        /// 更新控制模式下拉框选择
        /// </summary>
        private void UpdateControlModeSelection()
        {
            if (_controlModeSelectionUpdated)
                return;

            // 从配置表中获取工作模式地址配置
            if (TryGetConfig("工作模式", out var modeConfig))
            {
                // 解析地址
                if (TryParseAddress(modeConfig.AddressId, out var address))
                {
                    // 从缓存中获取值
                    string cachedValue = _modbusService?.GetCachedValue(address) ?? "";
                    if (!string.IsNullOrEmpty(cachedValue) && int.TryParse(cachedValue, out int index))
                    {
                        // 确保索引在有效范围内
                        if (index >= 0 && index < GroupOptions.Count)
                        {
                            // 更新下拉框选择
                            SelectedOption = GroupOptions[index];
                            _controlModeSelectionUpdated = true;
                        }
                    }
                }
            }
        }



        /// <summary>
        /// Modbus读取
        /// </summary>
        private async Task OptimizedReadDeviceData()
        {
            if (!_modbusService.IsConnected)
            {
                Debug.WriteLine($"[OptimizedRead] Modbus未连接");
                return;
            }

            try
            {
                var stopwatch = Stopwatch.StartNew();
                
                // 使用预加载的参数配置
                if (_preloadedReadParameters.Count == 0)
                {
                    Debug.WriteLine($"[OptimizedRead] 没有预加载的参数");
                    return;
                }

                // 计算需要读取的地址范围
                int minAddress = ParseAddress(_preloadedReadParameters.First().ParsID);
                int maxAddress = ParseAddress(_preloadedReadParameters.Last().ParsID);
                int count = maxAddress - minAddress + 1;
                
                Debug.WriteLine($"[OptimizedRead] 读取范围: {minAddress:X4}-{maxAddress:X4}, 数量: {count}");
                
                // 根据下拉框状态选择读取策略 
                if (!_hasChannelSelected)
                {
                    // 没有通道被选择：使用批量读取
                    var registers = await _modbusService.ReadMultipleHoldingRegistersAsync(1, minAddress.ToString("X4"), count);
                    if (registers != null && registers.Length > 0)
                    {
                        stopwatch.Stop();
                        Debug.WriteLine($"[OptimizedRead] 批量读取完成，处理 {registers.Length} 个寄存器，耗时: {stopwatch.ElapsedMilliseconds}ms");
                    }
                    else
                    {
                        Debug.WriteLine($"[OptimizedRead] 批量读取失败或返回空数据");
                    }
                }
                else
                {
                    // 有通道选择时，使用自定义报文读取
                    var registers = await _modbusService.ReadMultipleHoldingRegistersAsync1(1, minAddress.ToString("X4"), count);
                    if (registers != null && registers.Length > 0)
                    {
                        stopwatch.Stop();
                        Debug.WriteLine($"[OptimizedRead] 通道选择批量读取完成，处理 {registers.Length} 个寄存器，耗时: {stopwatch.ElapsedMilliseconds}ms");
                    }
                    else
                    {
                        Debug.WriteLine($"[OptimizedRead] 通道选择批量读取失败或返回空数据");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OptimizedRead] 读取异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行高频读取
        /// </summary>
        /// <param name="readMode">读取模式</param>
        /// <param name="readAll">是否读取所有数据（包括非"读"项）</param>
        private async Task ExecuteHighFrequencyRead(RealtimeReadMode readMode, bool readAll = false)
        {
            try
            {
                switch (readMode)
                {
                    case RealtimeReadMode.Serial:
                        await OptimizedReadDeviceData();
                        break;
                case RealtimeReadMode.Can:
                        await ExecuteHighFrequencyCanRead(readAll);
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[HighFrequencyRead] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// 执行高频CAN读取
        /// </summary>
        /// <param name="readAll">是否读取所有数据（包括非"读"项）</param>
        private async Task ExecuteHighFrequencyCanRead(bool readAll = false)
        {
            // 检查是否已停止高频读取
            if (!_isModbusReading)
            {
                return;
            }
            
            // 如果有事务正在进行，跳过本次读取，避免截获指令响应
            if (_otherConnectionService.TransactionInProgress)
            {
                Debug.WriteLine("[CanRead] 事务进行中，跳过读取");
                return;
            }
            
            if (!_otherConnectionService.IsStarted)
            {
                Debug.WriteLine("[CanRead] CAN未连接");
                return;
            }

            // 若已进入被动接收模式（示波器使能），由被动循环处理收到的数据，直接返回
            if (_isPassiveCanListening)
                return;

            if (!_hasChannelSelected)
            {
                try
                {
                    List<int> addresses;
                    
                    if (readAll)
                    {
                        // 读取所有数据（包括非"读"项）
                        addresses = _excelViewModel._allPers
                            .Select(p => ParseAddress(p.ParsID))
                            .Where(addr => addr > 0)
                            .Distinct()
                            .OrderBy(addr => addr)
                            .ToList();
                        Debug.WriteLine($"[CanRead] 读取所有数据，共 {addresses.Count} 个地址");
                    }
                    else
                    {
                        // 读取筛选后的数据（仅"读"项）
                        addresses = _preloadedReadParameters
                            .Select(p => ParseAddress(p.ParsID))
                            .Where(addr => addr > 0)
                            .Distinct()
                            .OrderBy(addr => addr)
                            .ToList();
                        Debug.WriteLine($"[CanRead] 读取筛选数据，共 {addresses.Count} 个地址");
                    }

                    if (addresses.Count == 0)
                    {
                        Debug.WriteLine("[CanRead] 没有有效的地址");
                        return;
                    }

                    // 将地址分割为连续的地址段
                    var addressSegments = new List<List<int>>();
                    var currentSegment = new List<int> { addresses[0] };

                    for (int i = 1; i < addresses.Count; i++)
                    {
                        if (addresses[i] == addresses[i - 1] + 1)
                        {
                            // 地址连续，添加到当前段
                            currentSegment.Add(addresses[i]);
                        }
                        else
                        {
                            // 地址不连续，结束当前段并开始新段
                            addressSegments.Add(currentSegment);
                            currentSegment = new List<int> { addresses[i] };
                        }
                    }
                    // 添加最后一个段
                    addressSegments.Add(currentSegment);

                    // 对每个连续地址段进行分批读取
                    foreach (var segment in addressSegments)
                    {
                        // 检查是否已停止高频读取
                        if (!_isModbusReading)
                        {
                            return;
                        }
                        
                        int startAddress = segment[0];
                        int endAddress = segment[^1];
                        int totalRegisters = endAddress - startAddress + 1;

                        // 分批读取，每次最多30个寄存器
                        int batchSize = 30;
                        int currentRegister = 0;

                        while (currentRegister < totalRegisters)
                        {
                            // 检查是否已停止高频读取
                            if (!_isModbusReading)
                            {
                                return;
                            }
                            
                            int registersToRead = Math.Min(batchSize, totalRegisters - currentRegister);
                            int currentStartAddress = startAddress + currentRegister;

                            // 构建读取报文，count参数为字节数，所以需要*2
                            var payload = _otherConnectionService.BuildRealtimeReadPayload(currentStartAddress, registersToRead * 2);
                            var response = await _otherConnectionService.SendRealtimeReadMessageAsync(payload);

                            if (response == null || response.Length == 0)
                            {
                                Debug.WriteLine($"[CanRead] 批量读取未收到响应或超时，地址: {currentStartAddress}, 数量: {registersToRead}");
                                currentRegister += registersToRead;
                                continue;
                            }

                            // 处理响应数据
                            if (response.Length >= 4 && response[0] == 0x60)
                            {
                                int parsedStart = (response[1] << 8) | response[2];
                                int dataLen = response[3];
                                if (dataLen >= 0 && response.Length >= 4 + dataLen)
                                {
                                    int registerCount = dataLen / 2;
                                    var values = new int[registerCount];
                                    for (int k = 0; k < registerCount; k++)
                                    {
                                        int hi = response[4 + k * 2];
                                        int lo = response[4 + k * 2 + 1];
                                        values[k] = (hi << 8) | lo;
                                    }
                                    _modbusService?.UpdateCacheFromExternal(parsedStart, values);
                                }
                            }
                            else if (response.Length >= 2)
                            {
                                var values = new int[response.Length / 2];
                                for (int k = 0; k < values.Length; k++)
                                {
                                    int hi = response[k * 2];
                                    int lo = response[k * 2 + 1];
                                    values[k] = (hi << 8) | lo;
                                }
                                _modbusService?.UpdateCacheFromExternal(currentStartAddress, values);
                            }

                            currentRegister += registersToRead;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[CanRead] 批量读取错误: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 启动高频Modbus读取
        /// </summary>
        public void StartHighFrequencyRead()
        {
            if (!_isModbusReading)
            {
                _isModbusReading = true;
                _modbusReadEvent.Set(); // 触发线程立即执行
                Debug.WriteLine($"[ModbusRead] 高频读取已启动，间隔: {_readIntervalMs}ms");
            }
        }

        /// <summary>
        /// 停止高频Modbus读取
        /// </summary>
        public void StopHighFrequencyRead()
        {
            _isModbusReading = false;
            Debug.WriteLine($"[ModbusRead] 高频读取已停止");
        }

        /// <summary>
        /// 解析十六进制地址字符串为整数
        /// </summary>
        private int ParseAddress(string addressHex)
        {
            if (int.TryParse(addressHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int address))
            {
                return address;
            }
            return 0;
        }

        /// <summary>
        /// 刷新预加载参数（当Excel配置变化时调用）
        /// </summary>
        public void RefreshPreloadedParameters()
        {
            PreloadReadParameters();
            _lastRegisterValues.Clear(); // 清空缓存，强制下次全部更新
            Debug.WriteLine($"[RefreshPreload] 预加载参数已刷新");
        }

        /// <summary>
        /// 设置读取间隔
        /// </summary>
        /// <param name="intervalMs">间隔毫秒数</param>
        public void SetReadInterval(int intervalMs)
        {
            if (intervalMs > 0 && intervalMs <= 1000)
            {
                _readIntervalMs = intervalMs;
                Debug.WriteLine($"[ModbusRead] 读取间隔设置为: {_readIntervalMs}ms");
            }
        }

        /// <summary>
        /// 清理专用线程资源
        /// </summary>
        private void CleanupModbusReadThread()
        {
            StopHighFrequencyRead();
            _modbusReadCts?.Cancel();
            _modbusReadEvent?.Set(); // 确保线程能退出等待
            
            if (_modbusReadThread?.IsAlive == true)
            {
                if (!_modbusReadThread.Join(1000)) // 等待1秒
                {
                    Debug.WriteLine($"[ModbusReadThread] 线程未能正常退出，强制终止");
                }
            }
            
            _modbusReadCts?.Dispose();
            _modbusReadEvent?.Dispose();
        }
        /// <summary>
        /// 读取配置表数据
        /// </summary>
        private void LoadJogConfigs()
        {
            _jogConfigs.Clear();
            foreach (var item in _jogConfigService.Load())
            {
                // 同名覆盖：确保以配置表最后一行为准
                if (!string.IsNullOrEmpty(item.Name))
                    _jogConfigs[item.Name] = item;
            }
        }

        /// <summary>
        /// 处理示波器开启/关闭控制
        /// </summary>
        /// <param name="hasChannelSelected">是否有通道被选择</param>
        private async Task HandleOscilloscopeControl(bool hasChannelSelected)
        {
            // 从配置表中获取示波器开启地址配置
            if (!TryGetConfig("示波器开启", out var oscilloscopeConfig))
                return;

            // 根据通道选择状态写入相应的值
            string valueToWrite = hasChannelSelected ? "1" : "0";
            var writeOk = await WriteConfigValueAsync(oscilloscopeConfig, valueToWrite);
            // 如果写入成功并且是开启，则启动被动接收循环
            if (writeOk)
            {
                if (valueToWrite == "1" && _otherConnectionService.IsStarted)
                {
                    StartPassiveCanListening();
                }
                else if (valueToWrite == "0")
                {
                    StopPassiveCanListening();
                }
            }
        }
         
        /// <summary>
        /// 处理通道选择索引写入逻辑
        /// </summary>
        private async Task HandleChannelSelectionIndex()
        {
            // 处理通道1选择
            await HandleChannelSelection(Channel1Selected, "通道一");
            
            // 处理通道2选择
            await HandleChannelSelection(Channel2Selected, "通道二");
            
            // 处理通道3选择
            await HandleChannelSelection(Channel3Selected, "通道三");
            
            // 处理通道4选择
            await HandleChannelSelection(Channel4Selected, "通道四");
        }

        /// <summary>
        /// 处理单个通道的选择逻辑
        /// </summary>
        /// <param name="selectedOption">选择的选项</param>
        /// <param name="configName">配置项名称</param>
        /// <param name="channelNumber">通道编号</param>
        private async Task HandleChannelSelection(string selectedOption, string configName)
        {
            // 从配置表中获取通道选择地址配置
            if (!TryGetConfig(configName, out var channelConfig))
                return;

            // 如果选择的是"无"，则写入0
            if (selectedOption == "无")
            {
                await WriteConfigValueAsync(channelConfig, "0");
                return;
            }

            // 获取选择项的索引
            int index = GetChannelOptionIndex(selectedOption);
            if (index >= 0)
            {
                await WriteConfigValueAsync(channelConfig, index.ToString());
            }
        }

        /// <summary>
        /// 获取通道选项在列表中的索引
        /// </summary>
        /// <param name="selectedOption">选择的选项</param>
        /// <returns>选项索引，如果未找到返回-1</returns>
        private int GetChannelOptionIndex(string selectedOption)
        {
            return ChannelOptions.IndexOf(selectedOption);
        }

        /// <summary>
        /// 处理控制模式选择，发送索引发送给设备
        /// </summary>
        private async Task HandleControlModeSelection()
        {
            // 从配置表中获取工作模式地址配置
            if (!TryGetConfig("工作模式", out var modeConfig))
                return;

            // 获取选择项的索引
            int index = GroupOptions.IndexOf(SelectedOption);
            if (index >= 0)
            {
                await WriteConfigValueAsync(modeConfig, index.ToString());
            }
        }

        /// <summary>
        /// 打开Excel参数表命令
        /// </summary>
        [RelayCommand]
        public void OpenExcel()
        {
            IsOpenExcel = true;
            var ExcelVm = _serviceProvider.GetRequiredService<ExcelViewModel>();
            _dialogService.ShowWindow<Excel, ExcelViewModel>(ExcelVm);
        }

        /// <summary>
        /// 打开串口连接命令
        /// </summary>
        [RelayCommand]
        public void OpenSerialPort()
        {
            var serialVm = _serviceProvider.GetRequiredService<SerialPortViewModel>();
            _dialogService.ShowWindow<Serialport, SerialPortViewModel>(serialVm);
        }

        /// <summary>
        /// 启动线程读取命令
        /// </summary>
        [RelayCommand]
        public void StartHighFrequencyReadCommand()
        {
            StartHighFrequencyRead();
        }

        /// <summary>
        /// 停止线程读取命令
        /// </summary>
        [RelayCommand]
        public void StopHighFrequencyReadCommand()
        {
            StopHighFrequencyRead();
        }

        /// <summary>
        /// 设置读取间隔命令
        /// </summary>
        /// <param name="intervalMs">间隔时间（毫秒）</param>
        [RelayCommand]
        public void SetReadIntervalCommand(string intervalMs)
        {
            if (int.TryParse(intervalMs, out int interval) && interval > 0)
            {
                SetReadInterval(interval);
            }
        }

        /// <summary>
        /// 打开其他连接（如CAN）命令
        /// </summary>
        [RelayCommand]
        private void OpenOtherConnection()
        {
            var otherConnectionVm = _serviceProvider.GetRequiredService<OtherConnectionViewModel>();
            _dialogService.ShowWindow<OtherConnection, OtherConnectionViewModel>(otherConnectionVm); 
        }

        /// <summary>
        /// 切换使能状态命令
        /// </summary>
        [RelayCommand]
        private async Task ToggleEnable()
        {
            // 尝试获取“使能”配置项
            if (!TryGetConfig("使能", out var enableConfig))
                return;

            if (!IsEnableLatched)
            {
                // 进入使能锁存：发送开状态后，点动按钮切为“单点模式”
                if (await WriteConfigValueAsync(enableConfig, enableConfig.OnState))
                {
                    IsEnableLatched = true;
                    EnableButtonText = "关使能";
                    ForwardJogButtonText = "正向单点";
                    ReverseJogButtonText = "反向单点";
                }
            }
            else
            {
                // 退出使能锁存：发送关状态并恢复“长按点动模式”
                if (await WriteConfigValueAsync(enableConfig, enableConfig.OffState))
                {
                    IsEnableLatched = false;
                    EnableButtonText = "使能";
                    ForwardJogButtonText = "正向点动";
                    ReverseJogButtonText = "反向点动";
                }
            }
        }

        [RelayCommand]
        private async Task Save()
        {
            // 尝试获取“保存”配置项
            if (!TryGetConfig("保存", out var enableConfig))
                return;

            if(await WriteConfigValueAsync(enableConfig, "6688"))
            {
                Console.WriteLine("保存成功");
            }

        }

        [RelayCommand]
        private async Task ClearFault()
        {
            // 尝试获取“清除故障”配置项
            if (!TryGetConfig("清除故障", out var faultConfig))
                return;

            if(await WriteConfigValueAsync(faultConfig, "0"))
            {
                Console.WriteLine("清除故障成功");
            }

        }

        /// <summary>
        /// 打开固件升级窗口命令
        /// </summary>
        [RelayCommand]
        public async Task OpenFirmwareUpdate()
        {
            // 停止读取操作
            StopHighFrequencyRead();
            
            // 停止CAN被动接收
            _otherConnectionService.StopPassiveReceive();
            
            var firmwareUpdateVm = _serviceProvider.GetRequiredService<FirmwareUpdateViewModel>();
            
            // 自定义显示窗口，以便添加关闭事件处理
            if (System.Windows.Application.Current?.Dispatcher != null && !System.Windows.Application.Current.Dispatcher.CheckAccess())
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() => OpenFirmwareUpdate());
                return;
            }

            // 1. 实例化窗口
            var window = new Views.FirmwareUpdate();

            // 2. 绑定 ViewModel
            window.DataContext = firmwareUpdateVm;

            System.Windows.Window? owner = System.Windows.Application.Current?.MainWindow;
            if (owner != null)
            {
                window.Owner = owner;
                window.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
            }
            else
            {
                window.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            }

            // 添加窗口关闭事件处理程序
            window.Closed += (sender, e) =>
            {
                // 窗口关闭时恢复读取操作
                StartHighFrequencyRead();
                
                // 恢复CAN被动接收
                _otherConnectionService.StartPassiveReceive();
            };

            // 使用非模态方式显示窗口
            window.Show();
            
            // 确保窗口不会抢占焦点
            window.Focusable = false;
        }

        #region
        private async Task WriteRegister(int address, int value)
        {
            byte[] data = new byte[2];
            data[0] = (byte)(value >> 8);
            data[1] = (byte)(value & 0xFF);
            var payload = _otherConnectionService.BuildWritePayload(address, data);
            Debug.WriteLine($"发送的数据: {BitConverter.ToString(payload)}");
            var result = _otherConnectionService.SendWriteMessage(payload);
            if (!result.Succeeded)
            {
                throw new System.Exception(result.Message);
            }
            await Task.Delay(50);
        }
        #endregion

        /// <summary>
        /// 提交当前给定值到设备
        /// </summary>
        /// <returns>操作是否成功</returns>
        public async Task<bool> CommitCurrentSetpointAsync()
        {
            // 手动编辑给定后回车提交：直接发送当前模式对应给定值
            return await SendSetpointAsync(false);
        }

        /// <summary>
        /// 提交参数值到设备
        /// </summary>
        /// <param name="parameterName">参数名称</param>
        /// <param name="value">参数值</param>
        /// <returns>操作是否成功</returns>
        public async Task<bool> CommitParameterValueAsync(string parameterName, string value)
        {
            if (!TryGetConfig(parameterName, out var config))
                return false;

            return await WriteConfigValueAsync(config, value);
        }



        /// <summary>
        /// 开始点动操作
        /// </summary>
        /// <param name="forward">是否为正向点动</param>
        public async Task StartJogAsync(bool forward)
        {
            if (IsEnableLatched)
            {
                // 锁存使能时，不需要“按下开/松开发关”，直接单次触发
                await JogSingleAsync(forward);
                return;
            }

            if (_isJogPressing)
                return;

            // 长按开始：先发送给定值（反向时发送负值）
            var setpointOk = await SendSetpointAsync(!forward);
            if (!setpointOk)
                return;
            if (!TryGetConfig("使能", out var enableConfig))
                return;

            // 再发送使能开状态，形成“点动开始”两条写指令
            var enableOk = await WriteConfigValueAsync(enableConfig, enableConfig.OnState);
            if (!enableOk)
                return;

            _isJogPressing = true;
            if (forward)
                ForwardJogButtonText = "正向点动中";
            else
                ReverseJogButtonText = "反向点动中";
        }

        /// <summary>
        /// 停止点动操作
        /// </summary>
        public async Task StopJogAsync()
        {
            if (!_isJogPressing || IsEnableLatched)
                return;
            if (!TryGetConfig("使能", out var enableConfig))
                return;

            // 长按松开：只发送使能关状态，结束点动
            await WriteConfigValueAsync(enableConfig, enableConfig.OffState);
            _isJogPressing = false;
            ForwardJogButtonText = "正向点动";
            ReverseJogButtonText = "反向点动";
        }

        /// <summary>
        /// 执行单点操作
        /// </summary>
        /// <param name="forward">是否为正向单点</param>
        public async Task JogSingleAsync(bool forward)
        {
            if (!IsEnableLatched)
                return;
            // 单点模式：每次点击都发送一次给定（反向取负）+ 一次使能开状态
            var setpointOk = await SendSetpointAsync(!forward);
            if (!setpointOk)
                return;
            if (!TryGetConfig("使能", out var enableConfig))
                return;
            await WriteConfigValueAsync(enableConfig, enableConfig.OnState);
        }

        /// <summary>
        /// 发送给定值到设备
        /// </summary>
        /// <param name="negative">是否发送负值（用于反向点动）</param>
        /// <returns>操作是否成功</returns>
        private async Task<bool> SendSetpointAsync(bool negative)
        {
            if (!TryGetConfig(_currentSetpointConfigName, out var setpointConfig))
                return false;

            // 读取界面输入值；反向点动仅在发送时取负，不回写页面显示
            var valueText = CurrentSetpointValue?.Trim() ?? string.Empty;
            
            // 处理电流给定值，需要乘以100
            if (_currentSetpointConfigName == "电流给定")
            {
                if (double.TryParse(valueText, out double currentValue))
                {
                    // 电流给定值乘以100
                    double scaledValue = currentValue * 100;
                    valueText = scaledValue.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    MessageBox.Show("电流给定的值格式错误。");
                    return false;
                }
            }
            
            if (negative)
            {
                if (!TryParseToInt(valueText, out var signed))
                {
                    MessageBox.Show($"{_currentSetpointConfigName} 的值格式错误。");
                    return false;
                }
                valueText = (-Math.Abs(signed)).ToString(CultureInfo.InvariantCulture);
            }

            return await WriteConfigValueAsync(setpointConfig, valueText);
        }

        /// <summary>
        /// 尝试获取配置项
        /// </summary>
        /// <param name="name">配置项名称</param>
        /// <param name="item">输出配置项</param>
        /// <returns>是否找到配置项</returns>
        private bool TryGetConfig(string name, out JogConfigItem item)
        {
            if (_jogConfigs.TryGetValue(name, out item!))
                return true;
            MessageBox.Show($"未找到配置项：{name}");
            return false;
        }

        /// <summary>
        /// 写入配置值到设备
        /// </summary>
        /// <param name="config">配置项</param>
        /// <param name="rawValue">原始值</param>
        /// <returns>操作是否成功</returns>
        private async Task<bool> WriteConfigValueAsync(JogConfigItem config, string rawValue)
        {
            if (!TryParseAddress(config.AddressId, out var address))
            {
                MessageBox.Show($"{config.Name} 地址无效：{config.AddressId}");
                return false;
            }   

            // 处理特殊参数
            string processedValue = rawValue;
            if (double.TryParse(rawValue, out double value))
            {
                switch (config.Name)
                {
                    case "电流比率给定":
                        // 电流比率给定：输入100以内的数据，当作百分比，传输给设备的值为最大电流*百分比
                        double maxCurrent = GetMaxCurrentFromCache();
                        double ratioValue = maxCurrent * (value / 100);
                        processedValue = ratioValue.ToString();
                        break;
                    case "电流有效值给定":
                        // 电流有效值给定：数据*1.414
                        double effectiveValue = value * 1.414;
                        processedValue = effectiveValue.ToString();
                        break;
                    case "电流峰值给定":
                        // 电流峰值给定：数据*100
                        double peakValue = value * 100;
                        processedValue = peakValue.ToString();
                        break;
                    case "旋转频率给定":
                        // 旋转频率给定：数据*10
                        double freqValue = value * 10;
                        processedValue = freqValue.ToString();
                        break;
                }
            }

            var writeMode = GetWriteProtocolMode();
            
            return writeMode switch
            {
                WriteProtocolMode.Serial => await WriteSerialConfigValueAsync(config, processedValue, address),
                WriteProtocolMode.Can => WriteCanConfigValue(config, processedValue, address),
                _ => ShowNoConnectionMessage()
                
            };
        }

        /// <summary>
        /// 通过串口写入配置值
        /// </summary>
        /// <param name="config">配置项</param>
        /// <param name="rawValue">原始值</param>
        /// <param name="address">地址</param>
        /// <returns>操作是否成功</returns>
        private async Task<bool> WriteSerialConfigValueAsync(JogConfigItem config, string rawValue, int address)
        {
            if (!string.IsNullOrEmpty(config.DataType) && config.DataType.Trim().ToLower() == "int32")
            {
                if (!TryParseToInt32(rawValue, out int int32Value))
                {
                    MessageBox.Show($"{config.Name} 值无效：{rawValue}，需要有效的32位整数");
                    return false;
                }

                // 处理32位整数，需要分高低位写入
                ushort highWord = (ushort)((int32Value >> 16) & 0xFFFF);
                ushort lowWord = (ushort)(int32Value & 0xFFFF);
                int previousAddress = address - 1;
                var values = new ushort[] { highWord, lowWord };
                var ok = await _modbusService.WriteMultipleRegistersAsync(1, previousAddress.ToString("X4"), values);
                if (!ok)
                    MessageBox.Show($"{config.Name} 批量写入失败。");
                return ok;
            }

            if (!TryParseToWordHex(rawValue, out var valueHex))
            {
                MessageBox.Show($"{config.Name} 值无效：{rawValue}");
                return false;
            }

            // 写入单个寄存器
            var singleWriteOk = await _modbusService.WriteSingleRegisterAsync(1, address.ToString("X4"), valueHex);
            if (!singleWriteOk)
                MessageBox.Show($"{config.Name} 写入失败。");
            return singleWriteOk;
        }

        /// <summary>
        /// 通过CAN写入配置值
        /// </summary>
        /// <param name="config">配置项</param>
        /// <param name="rawValue">原始值</param>
        /// <param name="address">地址</param>
        /// <returns>操作是否成功</returns>
        private bool WriteCanConfigValue(JogConfigItem config, string rawValue, int address)
        {
            if (!TryBuildCanWriteRequest(config, rawValue, address, out var startAddress, out var dataBytes, out var errorMessage))
            {
                MessageBox.Show(errorMessage);
                return false;
            }

            try
            {
                var payload = _otherConnectionService.BuildWritePayload(startAddress, dataBytes);
                var result = _otherConnectionService.SendWriteMessage(payload);
                if (!result.Succeeded)
                {
                    MessageBox.Show($"{config.Name} CAN写入失败。");
                    return false;
                }

                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show($"{config.Name} 写入数据过长。");
                return false;
            }
        }

        /// <summary>
        /// 构建CAN写入请求报文
        /// </summary>
        /// <param name="config">配置项</param>
        /// <param name="rawValue">原始值</param>
        /// <param name="address">地址</param>
        /// <param name="startAddress">输出起始地址</param>
        /// <param name="dataBytes">输出数据字节</param>
        /// <param name="errorMessage">输出错误信息</param>
        /// <returns>是否成功构建请求</returns>
        private bool TryBuildCanWriteRequest(JogConfigItem config, string rawValue, int address, out int startAddress, out byte[] dataBytes, out string errorMessage)
        {
            startAddress = address;
            dataBytes = Array.Empty<byte>();
            errorMessage = string.Empty;

            if (!string.IsNullOrEmpty(config.DataType) && config.DataType.Trim().ToLower() == "int32")
            {
                if (!TryParseToInt32(rawValue, out int int32Value))
                {
                    errorMessage = $"{config.Name} 值无效：{rawValue}，需要有效的32位整数";
                    return false;
                }

                // 处理32位整数，需要分高低位
                startAddress = address - 1;
                ushort highWord = (ushort)((int32Value >> 16) & 0xFFFF);
                ushort lowWord = (ushort)(int32Value & 0xFFFF);
                dataBytes = new byte[]
                {
                    (byte)(highWord >> 8),
                    (byte)(highWord & 0xFF),
                    (byte)(lowWord >> 8),
                    (byte)(lowWord & 0xFF)
                };
                return true;
            }

            if (!TryParseToWord(rawValue, out var word))
            {
                errorMessage = $"{config.Name} 值无效：{rawValue}";
                return false;
            }

            // 处理16位数据
            dataBytes = new byte[]
            {
                (byte)(word >> 8),
                (byte)(word & 0xFF)
            };
            return true;
        }

        /// <summary>
        /// 显示无连接消息
        /// </summary>
        /// <returns>总是返回false，表示操作失败</returns>
        private bool ShowNoConnectionMessage()
        {
            MessageBox.Show("未检测到可用的串口或CAN连接。");
            return false;
        }
        /// <summary>
        /// 将字符串转换为16进制对应的十进制数字
        /// </summary>
        /// <param name="text"></param>
        /// <param name="address"></param>
        /// <returns></returns>
        private static bool TryParseAddress(string text, out int address)
        {
            text = (text ?? string.Empty).Trim();
            if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                text = text[2..];
            if (!int.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out address))
                return false;
            return address >= 0 && address <= 0xFFFF;
        }

        /// <summary>
        /// 尝试将字符串解析为整数
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="value">输出整数值</param>
        /// <returns>是否解析成功</returns>
        private static bool TryParseToInt(string text, out int value)
        {
            text = (text ?? string.Empty).Trim();
            if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(text[2..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var hex))
                {
                    value = hex;
                    return true;
                }
            }
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return true;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var dv))
            {
                value = (int)Math.Round(dv);
                return true;
            }
            return false;
        }
         
        /// <summary>
        /// 尝试将字符串解析为16位无符号整数的十六进制表示
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="hex">输出十六进制字符串</param>
        /// <returns>是否解析成功</returns>
        private static bool TryParseToWordHex(string text, out string hex)
        {
            hex = string.Empty;
            if (!TryParseToInt(text, out var signed))
                return false;
            var word = signed & 0xFFFF;
            hex = word.ToString("X4");
            return true;
        }

        /// <summary>
        /// 尝试将字符串解析为16位无符号整数
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="word">输出16位无符号整数</param>
        /// <returns>是否解析成功</returns>
        private static bool TryParseToWord(string text, out ushort word)
        {
            word = 0;
            if (!TryParseToInt(text, out var signed))
                return false;
            word = (ushort)(signed & 0xFFFF);
            return true;
        }


        /// <summary>
        /// 将文字框中的文字转换为十进制整数
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="value">输出整数值</param>
        /// <returns>是否解析成功</returns>
        private static bool TryParseToInt32(string text, out int value)
        {
            value = 0;
            text = (text ?? string.Empty).Trim();
            
            
            // 支持10进制格式
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return true;
         
            
            return false;
        }

        /// <summary>
        /// 定时器 tick 事件处理
        /// </summary>
        private async void Timer_Tick(object? sender, EventArgs e)
        {
            // 从缓存更新UI界面
            UpdateUIFromCache();
        }

        /// <summary>
        /// 主窗体界面实时刷新
        /// </summary>
        /// <param name="parameterName">参数名称</param>
        /// <param name="value">参数值</param>
        private void UpdateViewModelProperty(string parameterName, string value)
        {
            
            
            // 确保在UI线程上更新属性
            Application.Current?.Dispatcher.Invoke(() =>
            {
                switch (parameterName)
                {
                    case "母线电压":
                        Voltage = value;
                        break;
                    case "驱动器温度":
                        DriveTemperature = value;
                        break;
                    case "电机温度":
                        MotorTemperature = value;
                        break;
                    case "电流给定":
                        CurrentSet = value+"A";
                        break;
                    case "电流反馈":
                        // 将电流反馈数据除100，保留一位小数
                        if (double.TryParse(value, out double currentValue))
                        {
                            double processedValue = currentValue / 100;
                            CurrentFeedbackPeak = processedValue.ToString("F1") + "A";
                        }
                        else
                        {
                            CurrentFeedbackPeak = value + "A";
                        }
                        break;
                    case "电流有效值":
                        CurrentRMS = value + "A";
                        break;
                    case "速度给定":
                        SpeedSet = value;
                        break;
                    case "速度反馈":
                        SpeedFeedback = value;
                        break;
                    case "位置给定":
                        CurrentSet = value;
                        PositionSet=value;
                        break;
                    case "位置反馈":
                        PositionFeedback = value;
                        break;
                }
            });
        }

        /// <summary>
        /// 停止数据采集
        /// </summary>
        public void StopDataCollection()
        {
            timer?.Stop();
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            timer?.Stop();
            CleanupModbusReadThread();
        }
        double i = 0;
        /// <summary>
        /// 处理CAN实时数据帧，提取示波器数据
        /// 设备每2ms发送一组数据，此方法负责解析并存入环形缓冲区
        /// </summary>
        /// <param name="payload">原始CAN数据载荷</param>
        /// <param name="channel">通道号 (1-4)</param>
        private void ProcessOscilloscopeData(byte[] payload, int channel)
        {
            try
            {
                // 解析数据点（跳过通道标识字节0x31-0x34）
                int idx = 2;
                var dataPoints = new List<double>();
                
                // 每2个字节解析为一个数据点（低字节在前，高字节在后）
                while (idx + 1 < payload.Length)
                {
                    int lo = payload[idx];
                    int hi = payload[idx + 1];
                    int rawValue = (hi << 8) | lo;
                    double value = ConvertRawValueToDouble(rawValue);

                    if (i==value-1)
                    {
                        i=value;
                    }
                    else{
                        Debug.WriteLine($"数据跳变: {i} -> {value}");
                        i =value;
                    }
                    dataPoints.Add(value);
                    idx += 2;
                }
                // 批量写入环形缓冲区
                if (dataPoints.Count > 0)
                {
                    _oscilloscopeBuffers[channel].Write(dataPoints.ToArray());

                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OscilloscopeData] 通道:{channel}:数据处理错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 从原始16位值转换为实际数值
        /// 处理有符号整数转换（16位有符号整数范围：-32768 到 32767）
        /// </summary>
        /// <param name="rawValue">原始16位整数值</param>
        /// <returns>转换后的实际数值</returns>
        private double ConvertRawValueToDouble(int rawValue)
        {
            // 处理16位有符号整数：如果值大于32767，则为负数
            if (rawValue > 32767) 
                rawValue -= 65536;
            return rawValue;
        }

        /// <summary>
        /// 获取指定通道的最新示波器数据
        /// 此方法用于定时器刷新时从缓存读取数据
        /// </summary>
        /// <param name="channel">通道号 (1-4)</param>
        /// <param name="pointCount">需要获取的数据点数量（默认100个点）</param>
        /// <returns>最新的数据点数组</returns>
        public double[] GetOscilloscopeData(int channel, int pointCount = 100)
        {
            lock (_oscilloscopeLocks[channel])
            {
                int readIndex = _readIndices[channel];
                var data = _oscilloscopeBuffers[channel].ReadFromIndex(ref readIndex, pointCount);
                _readIndices[channel] = readIndex;
                return data;
            }
        }
        
        /// <summary>
        /// 写入示波器数据到环形缓冲区（用于测试数据生成）
        /// </summary>
        /// <param name="channel">通道号 (1-4)</param>
        /// <param name="data">要写入的数据</param>
        public void WriteOscilloscopeData(int channel, double[] data)
        {
            if (data == null || data.Length == 0) return;
            
            lock (_oscilloscopeLocks[channel])
            {
                _oscilloscopeBuffers[channel].Write(data);
            }
        }

        /// <summary>
        /// 同步通道的读取索引，确保新通道从与其他通道相同的位置开始读取
        /// </summary>
        /// <param name="channel">通道号</param>
        public void SyncChannelReadIndex(int channel)
        {
            lock (_oscilloscopeLocks[channel])
            {
                // 找到当前读取索引最大的通道
                int maxReadIndex = _readIndices.Values.Max();
                // 将新通道的读取索引设置为最大索引
                _readIndices[channel] = maxReadIndex;
            }
        }

    }

    /// <summary>
    /// 数据点类，用于图表显示
    /// </summary>
    public class DataPoint
    {
        /// <summary>
        /// X坐标值
        /// </summary>
        public double X { get; set; }
        /// <summary>
        /// Y坐标值
        /// </summary>
        public double Y { get; set; }
    }
}
