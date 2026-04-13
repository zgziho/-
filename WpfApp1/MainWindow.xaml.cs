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
        
        /// <summary>
        /// 示波器刷新定时器 - 20Hz刷新率（50ms间隔）
        /// </summary>
        private System.Windows.Threading.DispatcherTimer _oscilloscopeTimer;
        
        /// <summary>
        /// 示波器是否正在运行
        /// </summary>
        private bool _isOscilloscopeRunning = false;
        
        /// <summary>
        /// 数据缓冲区 - 存储所有历史数据用于停止后查看
        /// </summary>
        private readonly Dictionary<int, List<double>> _dataBuffers = new()
        {
            { 1, new List<double>() },
            { 2, new List<double>() },
            { 3, new List<double>() },
            { 4, new List<double>() }
        };
        
        /// <summary>
        /// 当前显示的数据起始索引（用于停止后滑动查看）
        /// </summary>
        private int _displayStartIndex = 0;
        
        /// <summary>
        /// 每批次采集的数据点数
        /// </summary>
        private const int BATCH_POINT_COUNT = 500;
        
        /// <summary>
        /// 显示窗口大小
        /// </summary>
        private const int DISPLAY_WINDOW_SIZE = 500;
        
        /// <summary>
        /// 最大缓冲区大小（防止内存无限增长）
        /// </summary>
        private const int MAX_BUFFER_SIZE = 10000;
        
        /// <summary>
        /// 总共接收的数据点数（用于X轴无限增长）
        /// </summary>
        private long _totalPointCount = 0;
        
        /// <summary>
        /// 缓冲区起始点的全局索引（当缓冲区滚动时更新）
        /// </summary>
        private long _bufferStartIndex = 0;
        
        /// <summary>
        /// 触发水平值（上升沿触发电平）
        /// </summary>
        private double _triggerLevel = 0;
        
        /// <summary>
        /// 是否已触发（用于上升沿检测）
        /// </summary>
        private bool _hasTriggered = false;
        
        /// <summary>
        /// 上一次的数据值（用于上升沿检测）
        /// </summary>
        private double _lastTriggerValue = double.MinValue;
        
        /// <summary>
        /// 触发后的预触发数据缓冲区
        /// </summary>
        private List<double> _preTriggerBuffer = new List<double>();
        
        /// <summary>
        /// 预触发点数
        /// </summary>
        private const int PRE_TRIGGER_POINTS = 100;
        
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
        /// </summary>
        private void InitializeOscilloscope()
        {
            // 初始化ScottPlot控件显示设置
            InitializeScottPlot();
            
            // 创建示波器刷新定时器
            _oscilloscopeTimer = new System.Windows.Threading.DispatcherTimer();
            _oscilloscopeTimer.Interval = TimeSpan.FromMilliseconds(15);
            _oscilloscopeTimer.Tick += OscilloscopeTimer_Tick;
        }

        /// <summary>
        /// 初始化ScottPlot控件显示设置
        /// </summary>
        private void InitializeScottPlot()
        {
            if (MainPlot != null)
            {
                MainPlot.Plot.Clear();
                MainPlot.Plot.Axes.SetLimitsX(0, DISPLAY_WINDOW_SIZE);
                MainPlot.Plot.Legend.FontName = "微软雅黑";
                MainPlot.Plot.FigureBackground.Color = new Color("#F0F0F0");
                MainPlot.Plot.DataBackground.Color = new Color("#F0F0F0");
                MainPlot.Plot.Grid.MajorLineColor = new Color("#CCCCCC");
                MainPlot.Plot.Axes.Color(Colors.Black);
                MainPlot.Refresh();
            }
        }

        /// <summary>
        /// 示波器定时器触发事件
        /// </summary>
        private void OscilloscopeTimer_Tick(object? sender, EventArgs e)
        {
            if (!_isOscilloscopeRunning) return;
            
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
            for (int channel = 1; channel <= 4; channel++)
            {
                if (IsChannelSelected(viewModel, channel))
                {
                    // 获取最新数据
                    //var data = viewModel.GetOscilloscopeData(channel, BATCH_POINT_COUNT);

                    // 使用测试数据
                    var data = GenerateTestDataWithTrigger(channel, BATCH_POINT_COUNT);

                    // 将数据添加到缓冲区
                    AddDataToBuffer(channel, data);
                }
            }
            
            // 更新显示
            UpdatePlotDisplay();
        }

        /// <summary>
        /// 将数据添加到缓冲区
        /// </summary>
        private void AddDataToBuffer(int channel, double[] data)
        {
            var buffer = _dataBuffers[channel];
            buffer.AddRange(data);
            
            // 更新全局数据计数
            _totalPointCount += data.Length;
            
            // 限制缓冲区大小
            if (buffer.Count > MAX_BUFFER_SIZE)
            {
                int removeCount = buffer.Count - MAX_BUFFER_SIZE;
                buffer.RemoveRange(0, removeCount);
                
                // 更新缓冲区起始索引（X轴无限增长）
                _bufferStartIndex += removeCount;
            }
        }
        /// <summary>
        /// 更新图表显示
        /// </summary>
        private void UpdatePlotDisplay()
        {
            if (MainPlot == null) return;
            
            var viewModel = DataContext as MainWindowViewModel;
            if (viewModel == null) return;

            MainPlot.Plot.Clear();
            
            // 收集所有显示的数据点用于计算Y轴范围
            List<double> allDataPoints = new List<double>();
            
            // 为每个选中的通道绘制数据
            for (int channel = 1; channel <= 4; channel++)
            {
                if (IsChannelSelected(viewModel, channel))
                {
                    var buffer = _dataBuffers[channel];
                    if (buffer.Count == 0) continue;
                    
                    // 绘制完整的历史数据（使用全局索引，实现X轴无限增长）
                    double[] xData = new double[buffer.Count];
                    double[] yData = new double[buffer.Count];
                    
                    for (int i = 0; i < buffer.Count; i++)
                    {
                        xData[i] = _bufferStartIndex + i; // 使用全局索引
                        yData[i] = buffer[i];
                        allDataPoints.Add(buffer[i]);
                    }
                    
                    var scatter = MainPlot.Plot.Add.Scatter(xData, yData);
                    scatter.Color = GetChannelColor(channel);
                    scatter.LineWidth = 1;
                    scatter.MarkerSize = 0;
                }
            }
            
            // 计算并设置Y轴范围（添加10%的空白区域）
            if (allDataPoints.Count > 0)
            {
                double yMin = allDataPoints.Min();
                double yMax = allDataPoints.Max();
                double yRange = yMax - yMin;
                double margin = yRange * 0.1; // 10%的空白区域
                
                // 处理数据范围很小的情况
                if (yRange < 0.1)
                {
                    margin = 1.0; // 最小空白区域
                }
                
                MainPlot.Plot.Axes.SetLimitsY(yMin - margin, yMax + margin);
            }
            
            // 如果正在运行，自动跟随最新数据
            if (_isOscilloscopeRunning)
            {
                var firstBuffer = _dataBuffers.Values.FirstOrDefault(b => b.Count > 0);
                if (firstBuffer != null)
                {
                    // 使用全局索引设置X轴范围（实现X轴无限增长）
                    long endIndex = _bufferStartIndex + firstBuffer.Count;
                    long startIndex = Math.Max(_bufferStartIndex, endIndex - DISPLAY_WINDOW_SIZE);
                    MainPlot.Plot.Axes.SetLimitsX(startIndex, endIndex);
                }
            }
            else
            {
                // 停止模式下，保持当前视图不变，允许用户自由浏览
                // 不设置X轴范围，让ScottPlot的交互功能生效
            }
            
            MainPlot.Refresh();
        }

        /// <summary>
        /// 生成带上升沿触发的测试数据
        /// </summary>
        private double[] GenerateTestDataWithTrigger(int channel, int count)
        {
            double[] data = new double[count];
            Random random = new Random(channel * 123 + DateTime.Now.Millisecond);
            
            // 基础正弦波
            double phase = channel * Math.PI / 2;
            double frequency = 0.05 + channel * 0.02;
            
            for (int i = 0; i < count; i++)
            {
                // 生成基础信号
                double value = Math.Sin(i * frequency + phase) * 10;
                
                // 添加一些噪声
                value += random.NextDouble() * 2 - 1;
                
           
                
                data[i] = value;
            }
            
            return data;
        }

        /// <summary>
        /// 检查通道是否被选中
        /// </summary>
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
        /// 获取通道颜色
        /// </summary>
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

        /// <summary>
        /// 开始示波器
        /// </summary>
        private void StartOscilloscope()
        {
            _isOscilloscopeRunning = true;
            
            // 不需要清空缓冲区，保留历史数据，继续添加新数据
            
            _oscilloscopeTimer.Start();
        }

        /// <summary>
        /// 停止示波器
        /// </summary>
        private void StopOscilloscope()
        {
            _isOscilloscopeRunning = false;
            _oscilloscopeTimer.Stop();
            
            // 设置显示索引为最新数据位置，方便查看历史
            var firstBuffer = _dataBuffers.Values.FirstOrDefault(b => b.Count > 0);
            if (firstBuffer != null)
            {
                _displayStartIndex = Math.Max(0, firstBuffer.Count - DISPLAY_WINDOW_SIZE);
            }
        }

        /// <summary>
        /// 向左滑动查看历史数据
        /// </summary>
        private void ScrollLeft()
        {
            if (_isOscilloscopeRunning) return;
            
            _displayStartIndex = Math.Max(0, _displayStartIndex - DISPLAY_WINDOW_SIZE / 2);
            UpdatePlotDisplay();
        }

        /// <summary>
        /// 向右滑动查看历史数据
        /// </summary>
        private void ScrollRight()
        {
            if (_isOscilloscopeRunning) return;
            
            var firstBuffer = _dataBuffers.Values.FirstOrDefault(b => b.Count > 0);
            if (firstBuffer != null)
            {
                int maxStartIndex = Math.Max(0, firstBuffer.Count - DISPLAY_WINDOW_SIZE);
                _displayStartIndex = Math.Min(maxStartIndex, _displayStartIndex + DISPLAY_WINDOW_SIZE / 2);
            }
            UpdatePlotDisplay();
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
        /// <summary>
        /// 开始/停止示波器按钮点击事件
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (_isOscilloscopeRunning)
            {
                // 停止示波器
                StopOscilloscope();
                ((Button)sender).Content = "开始";
            }
            else
            {
                // 开始示波器
                StartOscilloscope();
                ((Button)sender).Content = "停止";
            }
        }
    }
}
