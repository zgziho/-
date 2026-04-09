using DocumentFormat.OpenXml.Drawing;
using Microsoft.Extensions.DependencyInjection;
using ScottPlot;
using ScottPlot.Plottables;
using ScottPlot.WPF;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {   
        private MainWindowViewModel viewModel;
        private bool _forwardPressed;
        private bool _reversePressed;

        private SpeechRecognitionEngine _recognizer;
        private SpeechSynthesizer _synthesizer;  // 语音合成引擎
        private bool _isAwake = false;  // 唤醒状态标志

        #region 示波器显示相关字段
        
       private DataLogger _dataLogger;

        private Dictionary<int, DataLogger> _channelDataLoggers = new Dictionary<int, DataLogger>();


        private DataStreamer _channel1;
        private DataStreamer _channel2;
        private DataStreamer _channel3;
        private DataStreamer _channel4;
        /// <summary>
        /// 示波器刷新定时器 - 20Hz刷新率（50ms间隔）
        /// 用于定时从ViewModel获取最新数据并更新ScottPlot显示
        /// </summary>
        private System.Windows.Threading.DispatcherTimer _oscilloscopeTimer;
        
        /// <summary>
        /// 示波器显示点数 - 固定500个点，实现一周期显示效果
        /// </summary>
        private const int OSCILLOSCOPE_POINT_COUNT = 500;
        
        /// <summary>
        /// 滚动缓冲区 - 用于存储历史数据，实现历史数据查看功能
        /// </summary>
        private readonly Dictionary<int, Queue<double>> _scrollBuffers = new()
        {
            { 1, new Queue<double>(10000) },
            { 2, new Queue<double>(10000) },
            { 3, new Queue<double>(10000) },
            { 4, new Queue<double>(10000) }
        };
        
        
        /// <summary>
        /// 显示窗口大小 - 固定显示500个点
        /// </summary>
        private const int DISPLAY_WINDOW_SIZE = 500;
        
        #endregion
        

        public MainWindow()
        {
            InitializeComponent();

            //语音识别功能
            Loaded += MainWindow_Loaded;

            var app = (App)Application.Current;
            if (app.ServiceProvider == null)
                throw new InvalidOperationException("ServiceProvider 未初始化。");

            viewModel = app.ServiceProvider.GetRequiredService<MainWindowViewModel>();
            DataContext = viewModel;

            Loaded += MainWindow_AutoOpenSerial;

            // 初始化示波器显示系统（新增）
            InitializeOscilloscope();


        }

        #region 示波器显示方法

        /// <summary>
        /// 初始化示波器显示系统
        /// 创建定时器并设置20Hz刷新率，初始化ScottPlot控件显示设置
        /// </summary>
        private void InitializeOscilloscope()
        {


            // 缓冲区大小 20000 点
            const int bufferSize = 20000;

            // 通道1 
            _channel1 = MainPlot.Plot.Add.DataStreamer(bufferSize);
            _channel1.Color = Color.FromHex("#1E90FF");
            _channel1.LineWidth = 2;
            _channel1.ViewScrollLeft(); 

            // 通道2 
            _channel2 = MainPlot.Plot.Add.DataStreamer(bufferSize);
            _channel2.Color = Color.FromHex("#FF4500");
            _channel2.LineWidth = 2;
            _channel2.ViewScrollLeft();

            // 通道3 
            _channel3 = MainPlot.Plot.Add.DataStreamer(bufferSize);
            _channel3.Color = Color.FromHex("#32CD32"); 
            _channel3.LineWidth = 2;
            _channel3.ViewScrollLeft();

            // 通道4 
            _channel4 = MainPlot.Plot.Add.DataStreamer(bufferSize);
            _channel4.Color = Color.FromHex("#FFA500"); 
            _channel4.LineWidth = 2;
            _channel4.ViewScrollLeft();



            // 创建示波器刷新定时器
            _oscilloscopeTimer = new System.Windows.Threading.DispatcherTimer();
            _oscilloscopeTimer.Interval = TimeSpan.FromMilliseconds(50);
            _oscilloscopeTimer.Tick += OscilloscopeTimer_Tick;
            _oscilloscopeTimer.Start();
            
            // 初始化ScottPlot控件显示设置
            InitializeScottPlot();
        }

        /// <summary>
        /// 初始化ScottPlot控件显示设置
        /// 设置标题、坐标轴标签等基础显示属性
        /// </summary>
        private void InitializeScottPlot()
        {
            if (MainPlot != null)
            {
                //_dataLogger = MainPlot.Plot.Add.DataLogger();
                //_dataLogger.Color = Colors.DodgerBlue;
                //_dataLogger.LineWidth = 2;

                ////_dataLogger.ViewSlide(width: 1000);
                //_dataLogger.ViewFull();

                // 显示图例
                MainPlot.Plot.Legend.FontName = "微软雅黑";
                //外层
                MainPlot.Plot.FigureBackground.Color = new Color("#F0F0F0");
                //数据背景颜色（内层）
                MainPlot.Plot.DataBackground.Color = new Color("#F0F0F0");
                //网格颜色
                MainPlot.Plot.Grid.MajorLineColor = new Color("#CCCCCC");
                //XY轴线颜色
                MainPlot.Plot.Axes.Color(Colors.Black);
                MainPlot.Refresh();
            }
        }

        /// <summary>
        /// 示波器定时器触发事件
        /// 每50ms执行一次，从ViewModel获取最新数据并刷新显示
        /// </summary>
        private void OscilloscopeTimer_Tick(object? sender, EventArgs e)
        {
            RefreshOscilloscopeDisplay();
        }

        /// <summary>
        /// 刷新示波器显示
        /// 获取所有选中通道的最新数据，并更新ScottPlot显示
        /// </summary>
        private void RefreshOscilloscopeDisplay()
        {
            if (MainPlot == null) return;
            
            var viewModel = DataContext as MainWindowViewModel;
            if (viewModel == null) return;

            // 获取所有选中通道的数据
            var channelData = new Dictionary<int, double[]>();
            
            for (int channel = 1; channel <= 4; channel++)
            {
                if (IsChannelSelected(viewModel, channel))
                {
                    // 从ViewModel获取最新数据
                    // var data = viewModel.GetOscilloscopeData(channel, OSCILLOSCOPE_POINT_COUNT);
                    
                    // 使用测试数据
                    var data = GenerateTestData(channel, OSCILLOSCOPE_POINT_COUNT);

                    switch (channel)
                    {
                        case 1: _channel1.AddRange(data); break;
                        case 2: _channel2.AddRange(data); break;
                        case 3: _channel3.AddRange(data); break;
                        case 4: _channel4.AddRange(data); break;

                    }



                    //if (data.Length > 0)
                    //{
                    //    // 将新数据添加到缓冲区
                    //    lock (_scrollBuffers[channel])
                    //    {
                    //        foreach (var point in data)
                    //        {
                    //            _scrollBuffers[channel].Enqueue(point);
                    //            // 保持缓冲区容量，超出部分从头部移除
                    //            while(_scrollBuffers[channel].Count > 9000)
                    //            {
                    //                _scrollBuffers[channel].Dequeue();
                    //            }
                    //        }
                    //    }
                        
                    //    // 获取当前显示窗口的数据
                    //    var scrollBuffer = _scrollBuffers[channel].ToArray();
                    //    int startIndex = Math.Max(0, scrollBuffer.Length - DISPLAY_WINDOW_SIZE );
                    //    int endIndex = Math.Min(scrollBuffer.Length , startIndex + DISPLAY_WINDOW_SIZE);
                        
                    //    if (startIndex < endIndex)
                    //    {
                    //        channelData[channel] = scrollBuffer[startIndex..endIndex];
                    //    }
                    //}
                }
            }
            MainPlot.Refresh();

            // 更新ScottPlot显示
            //UpdateScottPlotDisplay(channelData);
            
        }

        /// <summary>
        /// 测试数据生成方法 - 用于测试示波器显示
        /// </summary>
        /// <param name="channel">通道号</param>
        /// <param name="count">数据点数量</param>
        /// <returns>生成的测试数据</returns>
        private double[] GenerateTestData(int channel, int count)
        {
            double[] data = new double[count];
            Random random = new Random(channel * 123);
            double phase = channel * Math.PI / 2;
            double frequency = 0.02 + channel * 0.01;
            
            for (int i = 0; i < count; i++)
            {
                data[i] = Math.Sin(i * frequency + phase) * 10 + random.NextDouble() * 2 - 1;
            }
            
            return data;
        }

        /// <summary>
        /// 检查通道是否被选中
        /// </summary>
        /// <param name="viewModel">ViewModel实例</param>
        /// <param name="channel">通道号 (1-4)</param>
        /// <returns>是否选中</returns>
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

        /// <summary>
        /// 更新ScottPlot显示
        /// 实现固定X轴长度和滚动显示效果，支持历史数据查看
        /// </summary>
        /// <param name="channelData">通道数据字典</param>
        private void UpdateScottPlotDisplay(Dictionary<int, double[]> channelData)
        {
            bool channelsChanged = !_channelDataLoggers.Keys.SequenceEqual(channelData.Keys);
            
            if (channelsChanged)
            {
                MainPlot.Plot.Clear();
                _channelDataLoggers.Clear();

                foreach (var (channel, data) in channelData)
                {
                    var dataLogger = MainPlot.Plot.Add.DataLogger();
                    dataLogger.Color = GetChannelColor(channel);
                    dataLogger.LineWidth = 1;
                    
                    _channelDataLoggers[channel] = dataLogger;
                }
                
                if (channelData.Count > 1)
                {
                    MainPlot.Plot.ShowLegend();
                }
            }
            
            foreach (var (channel, data) in channelData)
            {
                if (data.Length > 0 && _channelDataLoggers.ContainsKey(channel))
                {
                    var dataLogger = _channelDataLoggers[channel];
                    
                    dataLogger.Clear();
                    
                    for (int i = 0; i < data.Length; i++)
                    {
                        dataLogger.Add((double)i, data[i]);
                    }
                }
            }
            
            if (channelData.Values.Any(d => d.Length > 0))
            {
                MainPlot.Plot.Axes.SetLimitsX(0, DISPLAY_WINDOW_SIZE - 1);
                
                var allData = channelData.Values.SelectMany(d => d).ToArray();
                if (allData.Length > 0)
                {
                    double yMin = allData.Min();
                    double yMax = allData.Max();
                    double yRange = yMax - yMin;
                    double margin = yRange * 0.1;
                    
                    MainPlot.Plot.Axes.SetLimitsY(yMin - margin, yMax + margin);
                }
            }
            
            MainPlot.Refresh();
        }

        /// <summary>
        /// 获取通道颜色
        /// 每个通道使用不同的颜色以便区分
        /// </summary>
        /// <param name="channel">通道号 (1-4)</param>
        /// <returns>对应的颜色</returns>
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

        #endregion

        private void MainWindow_AutoOpenSerial(object sender, RoutedEventArgs e)
        {
            // 仅首次加载自动弹出串口窗口，避免窗口重新激活时重复弹出
            Loaded -= MainWindow_AutoOpenSerial;
            viewModel.OpenSerialPortCommand.Execute(null);
        }
        /// <summary>
        /// 鼠标双击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetpointTextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                // 双击进入可编辑态
                textBox.IsReadOnly = false;
                textBox.Focus();
                textBox.SelectAll();
            }
        }
        /// <summary>
        /// 按下回车后
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void SetpointTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;
            if (sender is not TextBox textBox)
                return;

            // 回车时手动推送绑定源，再触发设备写入
            textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
            await viewModel.CommitCurrentSetpointAsync();
            textBox.IsReadOnly = true;
            Keyboard.ClearFocus();
            e.Handled = true;
        }
        /// <summary>
        /// 失去焦点后
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetpointTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            // 离开焦点后恢复只读，防止误触继续编辑
            if (sender is TextBox textBox)
                textBox.IsReadOnly = true;
        }

        /// <summary>
        /// 参数文本框双击事件：进入可编辑状态
        /// </summary>
        private void ParameterTextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                // 双击进入可编辑态
                textBox.IsReadOnly = false;
                textBox.Focus();
                textBox.SelectAll();
            }
        }

        /// <summary>
        /// 参数文本框按键事件：回车提交
        /// </summary>
        private async void ParameterTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;
            if (sender is not TextBox textBox)
                return;

            // 回车时提交参数值
            await CommitParameterValueAsync(textBox);
        }

        /// <summary>
        /// 参数文本框失去焦点事件：恢复只读状态
        /// </summary>
        private async void ParameterTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                // 失去焦点时提交参数值并恢复只读
                await CommitParameterValueAsync(textBox);
                textBox.IsReadOnly = true;
            }
        }

        /// <summary>
        /// 提交参数值到设备
        /// </summary>
        private async Task CommitParameterValueAsync(TextBox textBox)
        {
            string parameterName = textBox.Tag?.ToString() ?? string.Empty;
            string value = textBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(value))
                return;

            // 手动更新绑定源
            textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();

            // 调用ViewModel方法发送参数
            await viewModel.CommitParameterValueAsync(parameterName, value);
        }
        /// <summary>
        /// 左键按下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void ForwardJog_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (viewModel.IsEnableLatched)
                return;
            // 长按模式按下：触发“给定 + 使能开”
            _forwardPressed = true;
            await viewModel.StartJogAsync(true);
            e.Handled = true;
        }
        /// <summary>
        /// 左键松开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void ForwardJog_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_forwardPressed)
                return;
            // 长按模式松开：触发“使能关”
            _forwardPressed = false;
            await viewModel.StopJogAsync();
            e.Handled = true;
        }
        private async void ForwardJog_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!_forwardPressed)
                return;
            if (e.LeftButton == MouseButtonState.Pressed)
                return;
            _forwardPressed = false;
            await viewModel.StopJogAsync();
        }

        private async void ForwardJog_Click(object sender, RoutedEventArgs e)
        {
            if (!viewModel.IsEnableLatched)
                return;
            // 单点模式点击：一次性触发，不需要松开关断
            await viewModel.JogSingleAsync(true);
        }

        private async void ReverseJog_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (viewModel.IsEnableLatched)
                return;
            // 长按模式按下：反向点动（给定发送负值）
            _reversePressed = true;
            await viewModel.StartJogAsync(false);
            e.Handled = true;
        }

        private async void ReverseJog_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_reversePressed)
                return;
            // 长按模式松开：触发“使能关”
            _reversePressed = false;
            await viewModel.StopJogAsync();
            e.Handled = true;
        }

        private async void ReverseJog_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!_reversePressed)
                return;
            if (e.LeftButton == MouseButtonState.Pressed)
                return;
            _reversePressed = false;
            await viewModel.StopJogAsync();
        }

        private async void ReverseJog_Click(object sender, RoutedEventArgs e)
        {
            if (!viewModel.IsEnableLatched)
                return;
            // 单点模式点击：一次性触发，不需要松开关断
            await viewModel.JogSingleAsync(false);
        }

        /// <summary>
        /// 语音识别
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            //创建识别引擎
            _recognizer=new SpeechRecognitionEngine();
            //创建语音合成引擎
            _synthesizer = new SpeechSynthesizer();

            //定义唤醒词语法
            Choices wakeUpCommands = new Choices();
            wakeUpCommands.Add("你好语音助手","打开串口","打开参数表","关闭串口","关闭参数表");

            GrammarBuilder wakeUpGb = new GrammarBuilder(wakeUpCommands);
            Grammar wakeUpGrammar = new Grammar(wakeUpGb);
              
            //加载唤醒词语法
            _recognizer.LoadGrammar(wakeUpGrammar);
            //绑定识别成功事件
            _recognizer.SpeechRecognized += _recognizer_SpeechRecognized; 
            //设置输入设备并开始异步识别
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);

        }

        /// <summary>
        /// 判断语音操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _recognizer_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
        {
            string command = e.Result.Text;
            float confidence=e.Result.Confidence;

            if (confidence < 0.7) return;
            
            Dispatcher.Invoke(() =>
            {
                // 检查是否是唤醒词
                if (command == "语音助手")
                {
                    _isAwake = true;
                    // 使用语音播报代替对话框
                    _synthesizer.SpeakAsync("我在");
                    
                    // 5秒后自动退出唤醒状态
                    Task.Delay(5000).ContinueWith(_ => 
                    {
                        _isAwake = false;
                    });
                    return;
                }
                
                // 只有在唤醒状态下才执行其他指令
                if (!_isAwake)
                    return;
                
                switch (command)
                {
                    case "打开串口":
                        viewModel.OpenSerialPort();
                        break;
                    case "关闭串口":
                        CloseSerialPortWindow();
                        break;
                    case "打开参数表":
                        viewModel.OpenExcel();
                        break;
                    case "关闭参数表":
                        CloseExcelWindow();
                        break;
                    default: break;
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

        /// <summary>
        /// 关闭Excel窗口
        /// </summary>
        private void CloseExcelWindow()
        {
            // 查找当前打开的Excel窗口并关闭
            foreach (Window window in Application.Current.Windows)
            {
                if (window is WpfApp1.Views.Excel)
                {
                    window.Close();
                    break;
                }
            }
        }

        /// <summary>
        /// 关闭串口配置窗口
        /// </summary>
        private void CloseSerialPortWindow()
        {
            // 查找当前打开的串口配置窗口并关闭
            foreach (Window window in Application.Current.Windows)
            {
                if (window is WpfApp1.Views.Serialport)
                {
                    window.Close();
                    break;
                }
            }
        }
        bool ok= false;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (ok)
            {
                _oscilloscopeTimer.Tick += OscilloscopeTimer_Tick;
                _channel1.ManageAxisLimits = true;
                _channel2.ManageAxisLimits = true;
                _channel3.ManageAxisLimits = true;
                _channel4.ManageAxisLimits = true;
                ok = false;
            }
            else
            {
                _oscilloscopeTimer.Tick -= OscilloscopeTimer_Tick;
                _channel1.ManageAxisLimits = false;
                _channel2.ManageAxisLimits = false;
                _channel3.ManageAxisLimits = false;
                _channel4.ManageAxisLimits = false;
                ok = true;
            }
        }
    }
}
