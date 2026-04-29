using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using WpfApp1.Models;
using WpfApp1.Service;

namespace WpfApp1.ViewModels
{
    public partial class ExcelViewModel : ObservableObject
    {
        private ExcelService _excelService;
        public readonly ModbusService _modbusService;
        public readonly OtherConnectionService _otherConnectionService;
        private DispatcherTimer _continuousReadTimer;

        #region 绑定属性
        //切换下拉框
        [ObservableProperty]
        private ObservableCollection<string> _parList = new ObservableCollection<string>();
        //当前选中的文件
        [ObservableProperty]
        private string _document = string.Empty;
        // 数据集合
        [ObservableProperty]
        private ObservableCollection<Pars> _pers = new ObservableCollection<Pars>();

        // 当前选中的 Pars
        [ObservableProperty]
        private Pars _selectedPars = new Pars();

        [ObservableProperty]
        private Visibility _isProcessing = Visibility.Hidden;
        [ObservableProperty]
        private int _progressValue = 0;
        [ObservableProperty]
        private int _maxProgressValue = 100;

        // 用于新增/编辑的各字段
        [ObservableProperty]
        private string _newId = string.Empty;
        [ObservableProperty]
        private string _newParsID = string.Empty;
        [ObservableProperty]
        private string _newParsNM = string.Empty;
        [ObservableProperty]
        private string _newParsVA = string.Empty;
        [ObservableProperty]
        private string _newParsDW = string.Empty;
        [ObservableProperty]
        private string _newParsLX = string.Empty;
        [ObservableProperty]
        private string _newParsSX = string.Empty;
        [ObservableProperty]
        private string _newParsFW = string.Empty;
        [ObservableProperty]
        private string _newParsXSFS = string.Empty;
        [ObservableProperty]
        private string _newParsXS = string.Empty;
        [ObservableProperty]
        private string _newParsXSW = string.Empty;
        [ObservableProperty]
        private string _newParsLB = string.Empty;
        [ObservableProperty]
        private string _newParsYS = string.Empty;
        [ObservableProperty]
        private string _newParsBZ = string.Empty;

        // 搜索关键字（按参数名称或参数 ID 模糊搜索）
        [ObservableProperty]
        private string _searchText = string.Empty;

        // 当前选中的菜单项（1-5）
        [ObservableProperty]
        private int _selectedMenuIndex = 1;

        // 原始数据（用于筛选）
        public ObservableCollection<Pars> _allPers = new ObservableCollection<Pars>();

        #endregion
        public ExcelViewModel(ModbusService modbusService, OtherConnectionService otherConnectionService)
        {
            _modbusService = modbusService;
            _otherConnectionService = otherConnectionService;
            _excelService = new ExcelService();
            Scan();
            LoadData();
        }
        /// <summary>
        /// 扫描所有文件
        /// </summary>
        public void Scan()
        {
            try
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory+"..\\";
                var files = Directory.GetFiles(baseDir, "pars.xlsx");
                    foreach (string file in files) 
                {
                    ParList.Add(file);
                }             
            }
            catch (Exception ex)
            {
                Console.WriteLine("文件扫描错误"+ex.Message);
            }
            if (ParList.Count > 0)
            {
                Document = ParList[0].ToString();
                _excelService = new ExcelService(Document);
            }
        }
        /// <summary>
        /// 切换文件
        /// </summary>
        /// <param name="value">指定的文件</param>
        partial void OnDocumentChanged(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                _excelService = new ExcelService(value);
                LoadData();
            }
        }

        /// <summary>
        /// 加载数据
        /// </summary>
        private void LoadData()
        { 
            _allPers = _excelService.LoadData();
            ApplyFilter();
        }

        /// <summary>
        /// 应用筛选
        /// </summary>
        private void ApplyFilter()
        {
            if (_allPers == null || _allPers.Count == 0)
            {
                Pers = new ObservableCollection<Pars>();
                return;
            }

            // 根据 ParsLB 筛选（1-5）
            var filtered = _allPers.Where(p => 
            {
                if (string.IsNullOrEmpty(p.ParsLB)) return false;
                if (int.TryParse(p.ParsLB, out int lb))
                {
                    return lb == SelectedMenuIndex;
                }
                return false;
            }).ToList();

            Pers = new ObservableCollection<Pars>(filtered);
        }

        /// <summary>
        /// 切换菜单命令
        /// </summary>
        [RelayCommand]
        private void SelectMenu(int index)
        {
            SelectedMenuIndex = index;
            ApplyFilter();
        }

        /// <summary>
        /// 保存数据
        /// </summary>
        public  void SaveData()
        {
            _excelService.SaveData(_allPers);
        }
        /// <summary>
        /// 刷新命令
        /// </summary>
        [RelayCommand]
        private void Refresh()
        {
            LoadData();
        }

        /// <summary>
        /// 搜索命令（按 parsNM 或 parsID 模糊匹配）
        /// </summary>
        [RelayCommand]
        private void Search()
        {
            if (string.IsNullOrWhiteSpace(SearchText))
            {
                LoadData();
                return;
            }

            var filtered = Pers.Where(p =>
                (p.ParsNM?.Contains(SearchText) == true) ||
                (p.ParsID?.Contains(SearchText) == true)).ToList();
            Pers = new ObservableCollection<Pars>(filtered);
        }



        /// <summary>
        /// 另存为
        /// </summary>
        [RelayCommand]
        private void Download()
        {
            // 保存全部数据
            var allData = _allPers.ToList();

            if (!allData.Any())
            {
                MessageBox.Show("没有数据可以导出。", "提示",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 弹出保存文件对话框
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel 文件|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "pars.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    _excelService.SaveData(new ObservableCollection<Pars>(allData), saveFileDialog.FileName);

                    MessageBox.Show($"成功导出 {allData.Count} 条记录到：\n{saveFileDialog.FileName}",
                                    "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败：{ex.Message}", "错误",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }



 
    #region 实时读取相关配置
        /// <summary>
        /// 打开定时器实时读取
        /// </summary>
        public void StartAutoRead()
        {
            if (_continuousReadTimer == null)
            {
                _continuousReadTimer = new DispatcherTimer();
                _continuousReadTimer.Interval = TimeSpan.FromMilliseconds(100);
                _continuousReadTimer.Tick += ContinuousReadTimer_Tick;
            }
            _continuousReadTimer.Start();
        }
        /// <summary>
        /// 关闭实时读取
        /// </summary>
        public void StopAutoRead()
        {
            _continuousReadTimer?.Stop();
        }

        private async void ContinuousReadTimer_Tick(object? sender, EventArgs e)
        {
            await PerformAutoRead();
        }

         /// <summary>
        /// 找出连续的地址范围
        /// </summary>
        /// <param name="addresses"></param>
        /// <returns></returns>
        private List<(int startAddress, int count)> GetContiguousAddressRanges(List<int> addresses)
        {
            if (addresses.Count == 0) return new List<(int, int)>();

            var ranges = new List<(int startAddress, int count)>();
            addresses.Sort();

            int start = addresses[0];
            int count = 1;

            for (int i = 1; i < addresses.Count; i++)
            {
                if (addresses[i] == addresses[i-1] + 1) // 连续地址
                {
                    count++;
                }
                else // 非连续，开始新范围
                {
                    ranges.Add((start, count));
                    start = addresses[i];
                    count = 1;
                }
            }

            // 添加最后一个范围
            ranges.Add((start, count));

            return ranges;
        }
        #endregion

        /// <summary>
        /// 读取全部读写数据
        /// </summary>
        [RelayCommand]
        private async void ReadAsync()
        {
            // 检查至少有一种连接可用（串口或 CAN）
            if (!_modbusService.IsConnected && !_otherConnectionService.IsStarted)
            {
                MessageBox.Show("未连接任何设备，请先连接串口或启动 CAN 设备！");
                return;
            }

            // 检查是否有数据可读
            if (!Pers.Any())
            {
                MessageBox.Show("当前没有数据可以读取。");
                return;
            }
            try
            {
                _continuousReadTimer?.Stop();
                IsProcessing=Visibility.Visible;
                var readItems = Pers.Where(p => p.ParsSX != null && p.ParsSX.Trim() == "读写").ToList();
                if (readItems.Count > 0)
                {
                    await Display(readItems);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取过程中发生错误：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // 动画：确保在所有读取完成后显示 70% -> 停顿 -> 完成到 100%
                async Task AnimateCompletionAsync()
                {
                    try
                    {
                        if (MaxProgressValue <= 0)
                        {
                            // 无进度信息，简单延迟再隐藏
                            await Task.Delay(300).ConfigureAwait(true);
                            return;
                        }

                        var holdTarget = Math.Max(1, (int)Math.Floor(MaxProgressValue * 0.7));

                        // 如果当前进度低于 holdTarget，快速推进到 holdTarget
                        while (ProgressValue < holdTarget)
                        {
                            ProgressValue = Math.Min(holdTarget, ProgressValue + Math.Max(1, holdTarget / 10));
                            await Task.Delay(25).ConfigureAwait(true);
                        }

                        // 等待一段时间（确保用户感知完成停顿）
                        await Task.Delay(900).ConfigureAwait(true);

                        // 平滑填满到最终值
                        while (ProgressValue < MaxProgressValue)
                        {
                            ProgressValue = Math.Min(MaxProgressValue, ProgressValue + Math.Max(1, MaxProgressValue / 20));
                            await Task.Delay(20).ConfigureAwait(true);
                        }

                        // 保持短暂显示已满，然后隐藏
                        await Task.Delay(120).ConfigureAwait(true);
                    }
                    catch
                    {
                        // 忽略动画错误，继续清理界面状态
                    }
                }

                await AnimateCompletionAsync();
                IsProcessing = Visibility.Hidden;
                // 无论成功或失败都恢复定时器，避免自动读取停死
                ProgressValue = 0;
                _continuousReadTimer?.Start();
            }
        }
        /// <summary>
        /// 实时读取
        /// </summary>
        /// <returns></returns>
        private async Task PerformAutoRead()
        {
            if (!_modbusService.IsConnected && !_otherConnectionService.IsStarted)
                return;
            _continuousReadTimer?.Stop();
            try
            {
                // 获取所有parsSX为"读"的项
                var readItems = Pers.Where(p => p.ParsSX != null && p.ParsSX.Trim() == "读").ToList();

                if (readItems.Count > 0)
                    await DisplayFromCache(readItems);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自动读取过程中发生错误: {ex.Message}");
            }
            finally
            {
                _continuousReadTimer?.Start();
            }
        }
        /// <summary>
        /// 从内存缓存更新参数显示，不直接访问串口
        /// </summary>
        private async Task DisplayFromCache(List<Pars> readItems)
        {
            var registerAddresses = new List<int>();
            foreach (var item in readItems)
            {
                if (int.TryParse(item.ParsID, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int registerAddress))
                {
                    registerAddresses.Add(registerAddress);
                }
            }

            if (registerAddresses.Count == 0)
                return;

            var addressRanges = GetContiguousAddressRanges(registerAddresses);

            foreach (var (startAddress, count) in addressRanges)
            {
                // 仅接收时效内缓存，超过时效返回 null（防止展示过旧值）
                var registers = _modbusService.TryGetCachedHoldingRegisters(startAddress, count, 3000);
                if (registers == null || registers.Length == 0)
                    continue;
                for (int i = 0; i < registers.Length && i < count; i++)
                {
                    int currentAddress = startAddress + i;
                    var correspondingItem = readItems.FirstOrDefault(item =>
                        int.TryParse(item.ParsID, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int addr) &&
                        addr == currentAddress);

                    if (correspondingItem != null)
                    {
                        
                        // 检查数据类型，查找所需的高位值
                        int? prev1 = null, prev2 = null, prev3 = null;
                        string dataType = correspondingItem.ParsLX?.Trim().ToLower() ?? "";
                        
                        if (dataType == "int64")
                        {
                            // int64类型：查找前3行的值（高位到低位，当前行为最低位）
                            // prev1=地址-1(D2), prev2=地址-2(D1), prev3=地址-3(D0=最高位)
                            for (int j = 1; j <= 3; j++)
                            {
                                int prevIdx = i - j;
                                if (prevIdx >= 0 && prevIdx < registers.Length)
                                {
                                    if (j == 1) prev1 = registers[prevIdx];
                                    else if (j == 2) prev2 = registers[prevIdx];
                                    else if (j == 3) prev3 = registers[prevIdx];
                                }
                            }
                        }
                        else if (dataType == "int32")
                        {
                            // int32类型：查找上一行的值
                            int previousIndex = i - 1;
                            if (previousIndex >= 0 && previousIndex < registers.Length)
                            {
                                prev1 = registers[previousIndex];
                            }
                        }
                        
                        var processedValue = ProcessRegisterValue(registers[i], correspondingItem, prev1, prev2, prev3);
                        correspondingItem.ParsVA = processedValue;
                    }
                }
            }
            await Task.CompletedTask;
        }
       /// <summary>
       /// 读指令的实现逻辑
       /// </summary>
       /// <param name="readItems">需要查询的寄存器项</param>
       /// <returns></returns>
       private async Task Display(List<Pars> readItems)
        {
             // 解析所有寄存器地址
        var registerAddresses = new List<int>();
        foreach (var item in readItems)
        {
            if (int.TryParse(item.ParsID, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int registerAddress))
            {
                registerAddresses.Add(registerAddress);
            }
        }
        MaxProgressValue=registerAddresses.Count;

        if (registerAddresses.Count == 0)
            return;

        // 找出连续的地址范围
        var addressRanges = GetContiguousAddressRanges(registerAddresses);

        // 为每个连续地址范围执行批量读取（支持 Serial 或 CAN）
        foreach (var (startAddress, count) in addressRanges)
        {
            int[]? registers = null;

            // 优先使用串口 Modbus 读取
            if (_modbusService.IsConnected)
            {
                byte slaveAddress = 1; // 假设设备地址为1
                registers = await _modbusService.ReadMultipleHoldingRegistersAsync(slaveAddress, startAddress.ToString("X4"), count);
            }
            else if (_otherConnectionService.IsStarted)
            {
                // 使用 CAN 读取（OtherConnectionService 会执行写优先/互斥控制）
                try
                {
                    var canResult = await _otherConnectionService.SendRealtimeReadAndWaitForResponseAsync(startAddress, count);
                    registers = canResult;
                }
                catch
                {
                    registers = null;
                }
            }

            if (registers != null && registers.Length > 0)
            {
                // 将读取到的值分配给对应的项
                for (int i = 0; i < registers.Length && i < count; i++)
                {
                    int currentAddress = startAddress + i;
                    var correspondingItem = readItems.FirstOrDefault(item =>
                        int.TryParse(item.ParsID, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int addr) &&
                        addr == currentAddress);

                    if (correspondingItem != null)
                    {
                        // 在读取过程中不直接显示为 100%，保留一部分用于完成动画（例如保留到约 70%）
                        var holdTarget = Math.Max(1, (int)Math.Floor(MaxProgressValue * 0.7));
                        if (ProgressValue < holdTarget)
                        {
                            ProgressValue = Math.Min(holdTarget, ProgressValue + 1);
                        }

                        // 检查数据类型，查找所需的高位值
                        int? prev1 = null, prev2 = null, prev3 = null;
                        string dataType = correspondingItem.ParsLX?.Trim().ToLower() ?? "";
                        
                        if (dataType == "int64")
                        {
                            // int64类型：查找前3行的值（高位到低位，当前行为最低位）
                            // prev1=地址-1(D2), prev2=地址-2(D1), prev3=地址-3(D0=最高位)
                            for (int j = 1; j <= 3; j++)
                            {
                                int prevIdx = i - j;
                                if (prevIdx >= 0 && prevIdx < registers.Length)
                                {
                                    if (j == 1) prev1 = registers[prevIdx];
                                    else if (j == 2) prev2 = registers[prevIdx];
                                    else if (j == 3) prev3 = registers[prevIdx];
                                }
                            }
                        }
                        else if (dataType == "int32")
                        {
                            // int32类型：查找上一行的值
                            int previousIndex = i - 1;
                            if (previousIndex >= 0 && previousIndex < registers.Length)
                            {
                                prev1 = registers[previousIndex];
                            }
                        }

                        var processedValue = ProcessRegisterValue(registers[i], correspondingItem, prev1, prev2, prev3);
                        correspondingItem.ParsVA = processedValue;
                    }
                }
            }
            else
            {
                // 读取失败则退出当前操作
                return;
            }
        }
        }

        /// <summary>
        /// 处理显示的数据
        /// </summary>
        /// <param name="rawValue">原始寄存器值</param>
        /// <param name="parsItem">参数项</param>
        /// <param name="previousValue">上一行的寄存器值（用于int32类型组合）</param>
        /// <returns>处理后的显示值</returns>
        private string ProcessRegisterValue(int rawValue, Pars parsItem, int? prevValue1 = null, int? prevValue2 = null, int? prevValue3 = null)
        {
            double result = rawValue;
            
            // 检查数据类型并进行相应处理
            string dataType = parsItem.ParsLX?.Trim().ToLower() ?? string.Empty;
            
            switch (dataType)
            {
                case "int64":
                    // int64类型：递增方向为高位到低位
                    // 当前行为最低位(D3)，prevValue1=地址-1(D2), prevValue2=地址-2(D1), prevValue3=地址-3(D0)
                    // 组合：D0 << 48 | D1 << 32 | D2 << 16 | D3
                    if (prevValue1.HasValue && prevValue2.HasValue && prevValue3.HasValue)
                    {
                        long combinedValue = ((long)prevValue3.Value << 48) | 
                                            ((long)prevValue2.Value << 32) | 
                                            ((long)prevValue1.Value << 16) | 
                                            (rawValue & 0xFFFF);
                        result = combinedValue;
                    }
                    break;
                case "int32":
                    // int32类型：将当前行（低位）和上一行（高位）组合成32位整数
                    if (prevValue1.HasValue)
                    {
                        int combinedValue = (prevValue1.Value << 16) | (rawValue & 0xFFFF);
                        result = combinedValue;
                    }
                    break;
                case "int16":
                    // int16类型：将无符号16位值转换为有符号16位整数
                    // 原始值范围：0-65535，转换为int16范围：-32768到32767
                    result = (short)(rawValue & 0xFFFF);
                    break;
                case "uint16":
                    // uint16类型：无符号16位整数，直接使用原始值
                    // 范围：0-65535
                    result = rawValue & 0xFFFF;
                    break;
            }
            
            // 1. 与parsxs相乘（0-1范围内的小数）
            if (double.TryParse(parsItem.ParsXS, out double coefficient) && coefficient >= 0 && coefficient <= 1)
            {
                result *= coefficient;
            }
            
            // 2. 根据parsxsw确定小数位（F后面跟数字，表示保留几位小数）
            int decimalPlaces = 0;
            if (!string.IsNullOrEmpty(parsItem.ParsXSW) && parsItem.ParsXSW.StartsWith("F") && 
                int.TryParse(parsItem.ParsXSW.Substring(1), out int places))
            {
                decimalPlaces = places;
            }
            
            // 3. 根据parsxsfs确定显示格式
            string format = "D"; // 默认十进制
            if (!string.IsNullOrEmpty(parsItem.ParsXSFS))
            {
                if (parsItem.ParsXSFS.Trim().ToUpper() == "H")
                {
                    format = "X"; // 十六进制
                }
                else if (parsItem.ParsXSFS.Trim().ToUpper() == "D")
                {
                    format = "D"; // 十进制
                }
            }
            
            // 格式化输出
            if (format == "X")
            {
                // 十六进制显示，直接显示整数部分
                int intValue = (int)Math.Round(result);
                return intValue.ToString("X");
            }
            else
            {
                // 十进制显示，根据小数位格式化
                if (decimalPlaces > 0)
                {
                    return result.ToString($"F{decimalPlaces}");
                }
                else
                {
                    return ((int)Math.Round(result)).ToString();
                }
            }
        }

}}

    


       
