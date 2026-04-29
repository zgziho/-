using DocumentFormat.OpenXml.Drawing;
using Microsoft.Extensions.DependencyInjection;
using ScottPlot;
using ScottPlot.AxisLimitManagers;
using ScottPlot.Plottables;
using ScottPlot.WPF;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using System.Diagnostics;
using System.Data;
using WpfApp1.ViewModels;

namespace WpfApp1
{
    /// <summary>
    /// 触发采集状态枚举
    /// </summary>
    public enum AcquisitionState
    {
        /// <summary>
        /// 等待触发状态：持续更新预触发缓冲区，检测上升沿触发条件
        /// </summary>
        WaitingForTrigger,

        /// <summary>
        /// 捕获后触发状态：触发条件满足后，收集触发点之后的数据
        /// </summary>
        CapturingPostTrigger,

        /// <summary>
        /// Holdoff状态：采集完成后等待指定批次，防止抖动重复触发
        /// </summary>
        Holdoff
    }

    /// <summary>
    /// 单通道触发采集管理器
    /// <para>负责上升沿检测、预触发缓冲、跨批次数据拼接</para>
    /// <para>采用有限状态机模式管理采集流程，支持流式数据分批处理</para>
    /// </summary>
    public class TriggeredAcquisitionManager
    {
        #region 配置参数字段

        /// <summary>
        /// 预触发点数（触发点之前需要保存的样本数量）
        /// </summary>
        private readonly int _preTriggerPoints;

        /// <summary>
        /// 总显示窗口点数（预触发 + 后触发）
        /// </summary>
        private readonly int _totalDisplayPoints;

        /// <summary>
        /// 触发电平阈值（上升沿检测的参考电平）
        /// </summary>
        private double _triggerLevel;

        #endregion

        #region 状态管理字段

        /// <summary>
        /// 当前采集状态
        /// </summary>
        private AcquisitionState _state = AcquisitionState.WaitingForTrigger;

        /// <summary>
        /// 上一个采样值（用于上升沿检测的边缘比较）
        /// </summary>
        private double _lastSample = double.MinValue;

        #endregion

        #region 数据缓冲字段

        /// <summary>
        /// 预触发环形缓冲区：保存触发点之前的样本数据
        /// </summary>
        private readonly Queue<double> _preTriggerBuffer;

        /// <summary>
        /// 后触发缓冲区：保存触发点及之后的样本数据
        /// </summary>
        private readonly List<double> _postTriggerBuffer;

        /// <summary>
        /// 历史数据缓冲区：保存最近的采样数据，用于预触发缓冲区未满时填充
        /// </summary>
        private readonly Queue<double> _historyBuffer;

        /// <summary>
        /// 最大历史缓冲区大小（防止内存溢出）
        /// </summary>
        private const int MAX_HISTORY_SIZE = 10000;

        /// <summary>
        /// 还需采集的后触发样本数量
        /// </summary>
        private int _postTriggerNeeded;

        #endregion

        #region Holdoff 防抖字段

        /// <summary>
        /// Holdoff 计数器：剩余需要跳过的批次数量
        /// </summary>
        private int _holdoffCounter = 0;

        /// <summary>
        /// Holdoff 批次数量常量：每次采集完成后跳过的批次数
        /// </summary>
        private const int HOLDOFF_BATCHES = 2;

        #endregion

        #region 事件定义

        /// <summary>
        /// 采集完成事件：当完整的显示窗口数据捕获完成时触发
        /// </summary>
        /// <param name="capturedData">捕获的完整显示窗口数据（长度为 totalDisplayPoints）</param>
        public event Action<double[]>? CaptureCompleted;

        #endregion

        #region 属性

        /// <summary>
        /// 获取或设置触发电平阈值
        /// </summary>
        public double TriggerLevel
        {
            get => _triggerLevel;
            set => _triggerLevel = value;
        }

        #endregion

        #region 构造函数

        /// <summary>
        /// 初始化触发采集管理器
        /// </summary>
        /// <param name="preTriggerPoints">预触发点数（默认100）</param>
        /// <param name="totalDisplayPoints">总显示窗口点数（默认500）</param>
        /// <param name="triggerLevel">触发电平阈值（默认5.0）</param>
        public TriggeredAcquisitionManager(int preTriggerPoints = 100,
                                          int totalDisplayPoints = 500,
                                          double triggerLevel = 5.0)
        {
            _preTriggerPoints = preTriggerPoints;
            _totalDisplayPoints = totalDisplayPoints;
            _triggerLevel = triggerLevel;

            // 初始化缓冲区，预分配容量以提高性能
            _preTriggerBuffer = new Queue<double>(_preTriggerPoints + 1);
            _postTriggerBuffer = new List<double>(_totalDisplayPoints);
            _historyBuffer = new Queue<double>(MAX_HISTORY_SIZE);
        }

        #endregion

        #region 公共方法
        /// <summary>
        /// 处理批量采样数据
        /// </summary>
        /// <param name="samples">采样数据数组</param>
        public void ProcessDataBatch(double[] samples)
        {
            if (samples == null || samples.Length == 0)
                return;

            // 逐样本处理，确保每个样本都经过状态机逻辑
            for (int i = 0; i < samples.Length; i++)
            {
                double sample = samples[i];
                ProcessSingleSample(sample);
            }
        }
        /// <summary>
        /// 通知批次处理结束（用于Holdoff状态计数）
        /// </summary>
        public void NotifyBatchEnd()
        {
            if (_state == AcquisitionState.Holdoff)
            {
                _holdoffCounter--;
                if (_holdoffCounter <= 0)
                {
                    // Holdoff结束，恢复等待触发状态
                    _state = AcquisitionState.WaitingForTrigger;
                    // _preTriggerBuffer.Clear();
                    //_lastSample=double.MinValue;  
                }
            }
        }

        /// <summary>
        /// 重置采集管理器到初始状态
        /// </summary>
        public void Reset()
        {
            _state = AcquisitionState.WaitingForTrigger;
            _preTriggerBuffer.Clear();
            _postTriggerBuffer.Clear();
            _lastSample = double.MinValue;
            _holdoffCounter = 0;
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 处理单个采样点，根据当前状态分发到对应的处理逻辑
        /// </summary>
        /// <param name="sample">当前采样值</param>
        private void ProcessSingleSample(double sample)
        {
            switch (_state)
            {
                case AcquisitionState.WaitingForTrigger:
                    HandleWaitingState(sample);
                    break;
                case AcquisitionState.CapturingPostTrigger:
                    HandleCapturingState(sample);
                    break;
                case AcquisitionState.Holdoff:
                    UpdatePreTriggerBuffer(sample);  
                    break;
            }
            _lastSample = sample;  // 保存当前样本供下次边缘检测使用
        }

        /// <summary>
        /// 处理等待触发状态：检测上升沿，触发后切换到捕获状态
        /// </summary>
        /// <param name="sample">当前采样值</param>
        private void HandleWaitingState(double sample)
        {
            // 持续更新预触发缓冲区
            UpdatePreTriggerBuffer(sample);

            // 上升沿检测：上一个样本 < 阈值 && 当前样本 >= 阈值
            if (_lastSample < _triggerLevel && sample >= _triggerLevel)
            {
                // 触发条件满足，切换到捕获状态
                _state = AcquisitionState.CapturingPostTrigger;
                _postTriggerBuffer.Clear();


                // 计算还需采集的后触发点数，确保至少采集一定数量
                _postTriggerNeeded = _totalDisplayPoints - _preTriggerBuffer.Count;

                // 如果预触发缓冲已满，立即完成采集
                if (_postTriggerNeeded <= 0)
                    CompleteCapture();
            }
        }

        /// <summary>
        /// 处理捕获后触发状态：收集后触发数据，完成后触发采集完成事件
        /// </summary>
        /// <param name="sample">当前采样值</param>
        private void HandleCapturingState(double sample)
        {
            _postTriggerBuffer.Add(sample);
            _postTriggerNeeded--;

            // 后触发数据采集完成
            if (_postTriggerNeeded <= 0)
                CompleteCapture();
        }

        /// <summary>
        /// 更新预触发环形缓冲区
        /// </summary>
        /// <param name="sample">当前采样值</param>
        private void UpdatePreTriggerBuffer(double sample)
        {
            _preTriggerBuffer.Enqueue(sample);
            _historyBuffer.Enqueue(sample);
            
            // 保持预触发缓冲区大小恒定，超出时自动丢弃最旧的数据
            if (_preTriggerBuffer.Count > _preTriggerPoints)
                _preTriggerBuffer.Dequeue();
            
            // 保持历史缓冲区大小，超出时自动丢弃最旧的数据
            if (_historyBuffer.Count > MAX_HISTORY_SIZE)
                _historyBuffer.Dequeue();
        }

        /// <summary>
        /// 完成采集：拼接预触发和后触发数据，触发采集完成事件，进入Holdoff状态
        /// </summary>
        private void CompleteCapture()
        {
            // 创建完整的显示窗口
            double[] capturedWindow = new double[_totalDisplayPoints];
            int index = 0;

            // 确保预触发部分有 _preTriggerPoints 个点
            List<double> preTriggerData = new List<double>(_preTriggerBuffer);
            
            // 如果预触发缓冲区未满，从历史数据中填充
            if (preTriggerData.Count < _preTriggerPoints && _historyBuffer.Count > 0)
            {
                int needed = _preTriggerPoints - preTriggerData.Count;
                List<double> historyData = _historyBuffer.ToList();
                
                // 从历史数据中获取最旧的数据来填充
                for (int i = Math.Max(0, historyData.Count - needed); i < historyData.Count; i++)
                {
                    preTriggerData.Insert(0, historyData[i]);
                    if (preTriggerData.Count >= _preTriggerPoints)
                        break;
                }
            }
            
            // 确保预触发数据不超过 _preTriggerPoints
            if (preTriggerData.Count > _preTriggerPoints)
            {
                preTriggerData = preTriggerData.GetRange(preTriggerData.Count - _preTriggerPoints, _preTriggerPoints);
            }
            
            // 复制预触发数据
            foreach (var val in preTriggerData)
                if (index < _totalDisplayPoints)
                    capturedWindow[index++] = val;

            // 复制后触发数据
            foreach (var val in _postTriggerBuffer)
            {
                if (index < _totalDisplayPoints)
                    capturedWindow[index++] = val;
                else
                    break;
            }

            // 不足部分填0
            while (index < _totalDisplayPoints)
                capturedWindow[index++] = 0;

            // 触发事件通知外部
            CaptureCompleted?.Invoke(capturedWindow);

            // 重置状态
            _state = AcquisitionState.WaitingForTrigger;
            // 注意：不要清空预触发缓冲区，保持环形缓冲区的连续性
        }
        #endregion
    }

    public partial class MainWindow : Window
    {
        private MainWindowViewModel viewModel;
        private bool _forwardPressed;
        private bool _reversePressed;

        private SpeechRecognitionEngine _recognizer;
        private SpeechSynthesizer _synthesizer;
        private bool _isAwake = false;

        #region 示波器显示相关字段

        private System.Windows.Threading.DispatcherTimer _oscilloscopeTimer;
        private bool _isOscilloscopeRunning = false;

        /// <summary>
        /// 历史数据累积缓冲区（不清空，用于停止后滑动查看）
        /// </summary>
        private readonly Dictionary<int, List<double>> _historyBuffers = new()
        {
            { 1, new List<double>() },
            { 2, new List<double>() },
            { 3, new List<double>() },
            { 4, new List<double>() }
        };

        /// <summary>
        /// 缓冲区全局起始索引（用于X轴连续增长）
        /// </summary>
        private long _bufferStartIndex = 0;

        /// <summary>
        /// 最大历史缓冲区大小（防止内存溢出）
        /// </summary>
        private const int MAX_BUFFER_SIZE = 50000;

        private readonly Dictionary<int, TriggeredAcquisitionManager> _triggerManagers = new();

        private int _displayWindowSize = 300;
        private const int PRE_TRIGGER_POINTS = 500;
        private const double DEFAULT_TRIGGER_LEVEL = 5.0;

        /// <summary>
        /// 当前显示窗口大小（水平档位）
        /// </summary>
        public int DisplayWindowSize
        {
            get => _displayWindowSize;
            set
            {
                if (_displayWindowSize != value)
                {
                    _displayWindowSize = value;
                    ReinitializeTriggerManagers();
                    RefreshOscilloscopeDisplay();
                }
            }
        }

        /// <summary>
        /// 重新初始化触发采集管理器，使用新的显示窗口大小
        /// </summary>
        private void ReinitializeTriggerManagers()
        {
            _triggerManagers.Clear();
            for (int ch = 1; ch <= 4; ch++)
            {
                // 根据显示窗口大小计算合适的预触发区大小
                // 预触发区大小为显示窗口的1/3，最小30点，最大500点
                int preTriggerPoints = Math.Min(500, Math.Max(30, _displayWindowSize / 3));
                
                var manager = new TriggeredAcquisitionManager(
                    preTriggerPoints,
                    _displayWindowSize,
                    DEFAULT_TRIGGER_LEVEL);

                int channel = ch;
                manager.CaptureCompleted += (capturedData) =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (_isOscilloscopeRunning)
                        {
                            RefreshOscilloscopeDisplay(capturedData, channel);
                        }
                    });
                };

                _triggerManagers[ch] = manager;
            }
        }

        #region 测试数据生成器字段
        private System.Windows.Threading.DispatcherTimer? _testDataTimer;
        private long _testDataSampleCounter = 0;  // 全局样本计数器
        private readonly object _testDataLock = new object();
        #endregion

        #endregion

        public MainWindow()
        {
            InitializeComponent();

            //Loaded += MainWindow_Loaded;

            var app = (App)Application.Current;
            if (app.ServiceProvider == null)
                throw new InvalidOperationException("ServiceProvider 未初始化。");

            viewModel = app.ServiceProvider.GetRequiredService<MainWindowViewModel>();
            DataContext = viewModel;

            Loaded += MainWindow_AutoOpenSerial;

            InitializeOscilloscope();
        }

        #region 示波器显示方法

        private void InitializeOscilloscope()
        {
            InitializeScottPlot();

            for (int ch = 1; ch <= 4; ch++)
            {
                // 根据显示窗口大小计算合适的预触发区大小
                // 预触发区大小为显示窗口的1/3，最小30点，最大500点
                int preTriggerPoints = Math.Min(500, Math.Max(30, _displayWindowSize / 3));
                
                var manager = new TriggeredAcquisitionManager(
                    preTriggerPoints,
                    _displayWindowSize,
                    DEFAULT_TRIGGER_LEVEL);

                int channel = ch;
                manager.CaptureCompleted += (capturedData) =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (_isOscilloscopeRunning)
                        {
                            RefreshOscilloscopeDisplay(capturedData, channel);
                        }
                    });
                };

                _triggerManagers[ch] = manager;
            }

            _oscilloscopeTimer = new System.Windows.Threading.DispatcherTimer();
            _oscilloscopeTimer.Interval = TimeSpan.FromMilliseconds(5);
            _oscilloscopeTimer.Tick += OscilloscopeTimer_Tick;

            //#region 测试数据生成器 - 每毫秒生成50个数据写入环形缓冲区
            //_testDataTimer = new System.Windows.Threading.DispatcherTimer();
            //_testDataTimer.Interval = TimeSpan.FromMilliseconds(1);
            //_testDataTimer.Tick += TestDataTimer_Tick;
            //_testDataTimer.Start();
            //#endregion
        }

        private void InitializeScottPlot()
        {
            if (MainPlot != null)
            {
                MainPlot.Plot.Clear();
                MainPlot.Plot.Axes.SetLimitsX(0, _displayWindowSize);
                MainPlot.Plot.Legend.FontName = "微软雅黑";
                MainPlot.Plot.FigureBackground.Color = new Color("#F0F0F0");
                MainPlot.Plot.DataBackground.Color = new Color("#F0F0F0");
                MainPlot.Plot.Grid.MajorLineColor = new Color("#CCCCCC");
                MainPlot.Plot.Axes.Color(Colors.Black);
                MainPlot.Refresh();
            }
        }
        double y = 0;

        private void OscilloscopeTimer_Tick(object? sender, EventArgs e)
        {
            if (!_isOscilloscopeRunning) return;

            var vm = DataContext as MainWindowViewModel;
            if (vm == null) return;

            for (int channel = 1; channel <= 4; channel++)
            {
                if (IsChannelSelected(vm, channel))
                {//数据的获取
                    double[] data = viewModel.GetOscilloscopeData(channel, _displayWindowSize);

                    if (data.Length == 0) return;

                    _triggerManagers[channel].ProcessDataBatch(data);

                    // 直接将原生数据添加到历史缓冲区
                    var buffer = _historyBuffers[channel];
                    buffer.AddRange(data);
                    for (int i = 0; i < data.Length; i++)
                    {
                        //if (y == data[i] - 1)
                        //{
                        //    y = data[i];
                        //}
                        //else
                        //{
                        //    Debug.WriteLine($"数据变化：{y}===>{data[i]}");
                        //    y = data[i];
                        //}
                        if (y==data[i]-1||y == data[i] || y == data[i] + 1)
                        {
                            y = data[i];
                        }
                        else
                        {
                            if (y-data[i]>=500) {
                                Debug.WriteLine($"数据跳变: {y} ==> {data[i]}    |||跳动间隔==>   {y - data[i]}    |||批次号==>{data[1]}");

                            }
                            y = data[i];
                        }
                    }


                    // 确保所有通道缓冲区大小一致
                    SyncAllChannelBuffers();

                        // 限制缓冲区大小
                        if (buffer.Count > MAX_BUFFER_SIZE)
                        {
                            int removeCount = buffer.Count - MAX_BUFFER_SIZE;

                            // 所有通道同步删除相同数量的数据
                            foreach (var kvp in _historyBuffers)
                            {
                                kvp.Value.RemoveRange(0, removeCount);
                            }

                            _bufferStartIndex += removeCount;
                        }

                    
                }

                // 移除这里的刷新调用，因为数据捕获完成时会自动刷新
                // RefreshOscilloscopeDisplay();
            }
        }


        /// <summary>
        /// 快速刷新示波器（仅显示新数据）
        /// </summary>
        /// <param name="newData">新捕获的数据</param>
        /// <param name="channel">通道号</param>
        private void RefreshOscilloscopeDisplay(double[] newData, int channel)
        {
            if (MainPlot == null) return;
            var vm = DataContext as MainWindowViewModel;
            if (vm == null) return;

            MainPlot.Plot.Clear();
            
            // 只绘制当前通道的新数据
            if (IsChannelSelected(vm, channel))
            {
                // 计算新数据的索引范围
                long newDataStartIndex = _bufferStartIndex + (_historyBuffers[channel].Count - newData.Length);
                long newDataEndIndex = _bufferStartIndex + _historyBuffers[channel].Count - 1;
                
                double[] xData = new double[newData.Length];
                for (int i = 0; i < newData.Length; i++)
                    xData[i] = newDataStartIndex + i;
                
                var scatter = MainPlot.Plot.Add.Scatter(xData, newData);
                scatter.Color = GetChannelColor(channel);
                scatter.LineWidth = 1;
                scatter.MarkerSize = 0;
                
                // 设置Y轴范围
                double yMin = newData.Min();
                double yMax = newData.Max();
                double yRange = yMax - yMin;
                double margin = Math.Max(yRange * 0.1, 1.0);
                MainPlot.Plot.Axes.SetLimitsY(yMin - margin, yMax + margin);
                
                // 设置X轴范围
                MainPlot.Plot.Axes.SetLimitsX(newDataStartIndex, newDataEndIndex + 1);
            }
            
            MainPlot.Refresh();
        }

        /// <summary>
        /// 刷新示波器
        /// </summary>
        private void RefreshOscilloscopeDisplay()
        {
            if (MainPlot == null) return;
            var vm = DataContext as MainWindowViewModel;
            if (vm == null) return;

            MainPlot.Plot.Clear();
            List<double> allYValues = new List<double>();

            for (int channel = 1; channel <= 4; channel++)
            {
                if (IsChannelSelected(vm, channel))
                {
                    var buffer = _historyBuffers[channel];
                    if (buffer.Count == 0) continue;

                    double[] xData = new double[buffer.Count];
                    double[] yData = buffer.ToArray();
                    for (int i = 0; i < buffer.Count; i++)
                        xData[i] = _bufferStartIndex + i;
                    var scatter = MainPlot.Plot.Add.Scatter(xData, yData);
                    scatter.Color = GetChannelColor(channel);
                    scatter.LineWidth = 1;
                    scatter.MarkerSize = 0;
                    allYValues.AddRange(yData);
                }
            }

            if (allYValues.Count > 0)
            {
                double yMin = allYValues.Min();
                double yMax = allYValues.Max();
                double yRange = yMax - yMin;
                double margin = Math.Max(yRange * 0.1, 1.0);
                MainPlot.Plot.Axes.SetLimitsY(yMin - margin, yMax + margin);
            }

            // 运行状态：X轴跟随最新数据；停止状态：不自动设置范围，保留用户交互视图
            if (_isOscilloscopeRunning)
            {
                var firstBuffer = _historyBuffers.Values.FirstOrDefault(b => b.Count > 0);
                if (firstBuffer != null)
                {
                    long endIndex = _bufferStartIndex + firstBuffer.Count;
                    long startIndex = Math.Max(_bufferStartIndex, endIndex - _displayWindowSize);
                    MainPlot.Plot.Axes.SetLimitsX(startIndex, endIndex);
                }
            }
            // 停止时完全不设置X轴范围，让ScottPlot保持当前视图
            //Transparent
            MainPlot.Refresh();
        }

        /// <summary>
        /// 同步所有通道的缓冲区大小，确保它们始终保持一致
        /// 新通道会自动填充NaN对齐到当前最大缓冲区大小
        /// </summary>
        private void SyncAllChannelBuffers()
        {
            // 找到最大的缓冲区大小
            int maxBufferSize = _historyBuffers.Values.Max(b => b.Count);
            
            // 确保所有通道的缓冲区大小一致
            foreach (var kvp in _historyBuffers)
            {
                var chBuffer = kvp.Value;
                int currentSize = chBuffer.Count;
                
                // 如果当前通道缓冲区小于最大大小，填充NaN
                if (currentSize < maxBufferSize)
                {
                    int paddingCount = maxBufferSize - currentSize;
                    chBuffer.InsertRange(0, Enumerable.Repeat(double.NaN, paddingCount));
                }
            }
        }

        //// 所有通道共用一个随机数生成器（固定种子，确保噪声模式一致）
        //private readonly Random _sharedRandom = new Random(12345);

        //#region 测试数据生成器 - 每毫秒生成50个数据写入环形缓冲区
        //private void TestDataTimer_Tick(object? sender, EventArgs e)
        //{
        //    lock (_testDataLock)
        //    {
        //        const int POINTS_PER_MS = 50;  // 每毫秒生成50个数据点
        //        long currentCounter = _testDataSampleCounter;

        //        for (int channel = 1; channel <= 4; channel++)
        //        {
        //            double[] data = new double[POINTS_PER_MS];
        //            double frequency = 0.008;  // 较低频率，方波更宽
        //            double amplitude = 15;      // 更大幅度，正负明显
        //            double phase = (channel - 1) * Math.PI / 4;  // 不同通道相位偏移

        //            for (int i = 0; i < POINTS_PER_MS; i++)
        //            {
        //                long sampleIndex = currentCounter + i;
        //                // 方波生成：正弦值>0时为高电平，否则为低电平
        //                double rawValue = Math.Sin(2 * Math.PI * sampleIndex * frequency + phase);
        //                double squareWave = rawValue > 0 ? amplitude : -amplitude;
        //                // 添加少量噪声模拟真实信号
        //                squareWave += _sharedRandom.NextDouble() * 0.3 - 0.15;
        //                data[i] = squareWave;
        //            }

        //            // 写入到 ViewModel 的环形缓冲区
        //            var vm = DataContext as MainWindowViewModel;
        //            vm?.WriteOscilloscopeData(channel, data);
        //        }

        //        _testDataSampleCounter += POINTS_PER_MS;
        //    }
        //}
        //#endregion

        

        private bool IsChannelSelected(MainWindowViewModel viewModel, int channel)
        {
            return channel switch
            {
                1 => viewModel.Channel1Selected != "无",
                2 => viewModel.Channel2Selected != "无",
                3 => viewModel.Channel3Selected != "无",
                4 => viewModel.Channel4Selected != "无",
                _ => false
            };
        }

        private ScottPlot.Color GetChannelColor(int channel)
        {
            return channel switch
            {
                1 => ScottPlot.Colors.Blue,
                2 => ScottPlot.Colors.Red,
                3 => ScottPlot.Colors.Green,
                4 => ScottPlot.Colors.Orange,
                _ => ScottPlot.Colors.Black
            };
        }

        private void StartOscilloscope()
        {
            _isOscilloscopeRunning = true;

            foreach (var manager in _triggerManagers.Values)
                manager.Reset();

            _oscilloscopeTimer.Start();
        }

        private void StopOscilloscope()
        {
            _isOscilloscopeRunning = false;
            _oscilloscopeTimer.Stop();
            // 停止时刷新完整历史数据
            RefreshOscilloscopeDisplay();
        }

        #endregion

        private void MainWindow_AutoOpenSerial(object sender, RoutedEventArgs e)
        {
            Loaded -= MainWindow_AutoOpenSerial;
            viewModel.OpenSerialPortCommand.Execute(null);
        }

        private void SetpointTextBox_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.IsReadOnly = false;
                textBox.Focus();
                textBox.SelectAll();
                e.Handled = true;
            }
        }

        private void SetpointTextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.IsReadOnly = false;
                textBox.Focus();
                textBox.SelectAll();
            }
        }

        private async void SetpointTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            if (sender is not TextBox textBox) return;

            textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            await viewModel.CommitCurrentSetpointAsync();
            textBox.IsReadOnly = true;
            Keyboard.ClearFocus();
            e.Handled = true;
        }

        private void SetpointTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox)
                textBox.IsReadOnly = true;
        }

        private void ParameterTextBox_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.IsReadOnly = false;
                textBox.Focus();
                textBox.SelectAll();
                e.Handled = true;
            }
        }

        private void ParameterTextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.IsReadOnly = false;
                textBox.Focus();
                textBox.SelectAll();
            }
        }

        private async void ParameterTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            if (sender is not TextBox textBox) return;
            await CommitParameterValueAsync(textBox);
        }

        private async void ParameterTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                await CommitParameterValueAsync(textBox);
                textBox.IsReadOnly = true;
            }
        }

        private async Task CommitParameterValueAsync(TextBox textBox)
        {
            string parameterName = textBox.Tag?.ToString() ?? string.Empty;
            string value = textBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(value))
                return;

            textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            await viewModel.CommitParameterValueAsync(parameterName, value);
        }

        private async void ForwardJog_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (viewModel.IsEnableLatched) return;
            _forwardPressed = true;
            await viewModel.StartJogAsync(true);
            e.Handled = true;
        }

        private async void ForwardJog_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_forwardPressed) return;
            _forwardPressed = false;
            await viewModel.StopJogAsync();
            e.Handled = true;
        }

        private async void ForwardJog_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!_forwardPressed) return;
            if (e.LeftButton == MouseButtonState.Pressed) return;
            _forwardPressed = false;
            await viewModel.StopJogAsync();
        }

        private async void ForwardJog_Click(object sender, RoutedEventArgs e)
        {
            if (!viewModel.IsEnableLatched) return;
            await viewModel.JogSingleAsync(true);
        }
                                          
        private async void ReverseJog_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (viewModel.IsEnableLatched) return;
            _reversePressed = true;
            await viewModel.StartJogAsync(false);
            e.Handled = true;
        }

        private async void ReverseJog_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_reversePressed) return;
            _reversePressed = false;
            await viewModel.StopJogAsync();
            e.Handled = true;
        }

        private async void ReverseJog_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!_reversePressed) return;
            if (e.LeftButton == MouseButtonState.Pressed) return;
            _reversePressed = false;
            await viewModel.StopJogAsync();
        }

        private async void ReverseJog_Click(object sender, RoutedEventArgs e)
        {
            if (!viewModel.IsEnableLatched) return;
            await viewModel.JogSingleAsync(false);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _recognizer = new SpeechRecognitionEngine();
            _synthesizer = new SpeechSynthesizer();

            Choices wakeUpCommands = new Choices();
            wakeUpCommands.Add("你好语音助手", "打开串口", "打开参数表", "关闭串口", "关闭参数表");

            GrammarBuilder wakeUpGb = new GrammarBuilder(wakeUpCommands);
            Grammar wakeUpGrammar = new Grammar(wakeUpGb);

            _recognizer.LoadGrammar(wakeUpGrammar);
            _recognizer.SpeechRecognized += _recognizer_SpeechRecognized;
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void _recognizer_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
        {
            string command = e.Result.Text;
            float confidence = e.Result.Confidence;
            if (confidence < 0.7) return;

            Dispatcher.Invoke(() =>
            {
                if (command == "语音助手")
                {
                    _isAwake = true;
                    _synthesizer.SpeakAsync("我在");
                    Task.Delay(5000).ContinueWith(_ => _isAwake = false);
                    return;
                }
                if (!_isAwake) return;

                switch (command)
                {
                    case "打开串口": viewModel.OpenSerialPort(); break;
                    case "关闭串口": CloseSerialPortWindow(); break;
                    case "打开参数表": viewModel.OpenExcel(); break;
                    case "关闭参数表": CloseExcelWindow(); break;
                }
            });
        }

        protected override void OnClosed(EventArgs e)
        {
            _recognizer?.RecognizeAsyncStop();
            _recognizer?.Dispose();
            _synthesizer?.Dispose();
            base.OnClosed(e);
        }

        private void CloseExcelWindow()
        {
            foreach (Window window in Application.Current.Windows)
                if (window is WpfApp1.Views.Excel)
                    window.Close();
        }

        private void CloseSerialPortWindow()
        {
            foreach (Window window in Application.Current.Windows)
                if (window is WpfApp1.Views.Serialport)
                    window.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (_isOscilloscopeRunning)
            {
                StopOscilloscope();
                ((Button)sender).Content = "开始";
            }
            else
            {
                StartOscilloscope();
                ((Button)sender).Content = "停止";
            }
        }

        private void HorizontalRangeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox comboBox && comboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                if (int.TryParse(selectedItem.Tag?.ToString(), out int newWindowSize))
                {
                    DisplayWindowSize = newWindowSize;
                }
            }
        }
    }
}