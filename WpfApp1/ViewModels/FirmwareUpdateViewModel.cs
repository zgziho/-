using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using WpfApp1.Service;
using ZLGCAN;

namespace WpfApp1.ViewModels
{
    /// <summary>
    /// 固件升级视图模型
    /// </summary>
    public partial class FirmwareUpdateViewModel : ObservableObject
    {
        /// <summary>
        /// Modbus服务，用于串口通信
        /// </summary>
        private readonly ModbusService _modbusService;
        
        /// <summary>
        /// 其他连接服务，用于CAN通信
        /// </summary>
        private readonly OtherConnectionService _otherConnectionService;

        /// <summary>
        /// 发送次数
        /// </summary>
        private int _packetCount = 0;

        /// <summary>
        /// 固件文件路径
        /// </summary>
        [ObservableProperty]
        private string firmwareFilePath = string.Empty;

        /// <summary>
        /// 是否本地升级
        /// </summary>
        [ObservableProperty]
        private bool isLocalUpdate = true;

        /// <summary>
        /// 是否远程升级
        /// </summary>
        [ObservableProperty]
        private bool isRemoteUpdate = false;

        /// <summary>
        /// 设备信息
        /// </summary>
        [ObservableProperty]
        private string deviceInfo = "未连接设备";

        /// <summary>
        /// 升级进度值
        /// </summary>
        [ObservableProperty]
        private int progressValue = 0;

        /// <summary>
        /// 进度信息
        /// </summary>
        [ObservableProperty]
        private string progressInfo = string.Empty;

        /// <summary>
        /// 状态消息
        /// </summary>
        [ObservableProperty]
        private string statusMessage = "就绪";
        
        /// <summary>
        /// 升级是否完成
        /// </summary>
        private bool _updateCompleted = false;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="modbusService">Modbus服务</param>
        /// <param name="otherConnectionService">其他连接服务</param>
        public FirmwareUpdateViewModel(ModbusService modbusService, OtherConnectionService otherConnectionService)
        {
            _modbusService = modbusService;
            _otherConnectionService = otherConnectionService;
        }

        /// <summary>
        /// 浏览固件文件命令
        /// </summary>
        [RelayCommand]
        private void Browse()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "固件文件 (*.bin;*.hex)|*.bin;*.hex|所有文件 (*.*)|*.*";
            openFileDialog.Title = "选择固件文件";

            if (openFileDialog.ShowDialog() == true)
            {
                FirmwareFilePath = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// 开始升级命令
        /// </summary>
        [RelayCommand]
        private async Task StartUpdate()
        {
            //if (string.IsNullOrEmpty(FirmwareFilePath))
            //{
            //    MessageBox.Show("请选择固件文件", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            StatusMessage = "升级中...";
            ProgressValue = 0;
            ProgressInfo = "开始升级...";

            try
            {
                //步骤1：向地址0x10CE写0x6688，等待3s
                ProgressInfo = "准备进入升级模式";
                await WriteRegister(0x10CE, 0x6688);
                await Task.Delay(3000);

                // 步骤2：向地址0x9002写十个寄存器值为0，等待100ms
                ProgressInfo = "初始化升级参数";
                byte[] zeroData = new byte[10]; 
                await WriteMultipleRegisters(0x9002, zeroData);
                await Task.Delay(200);

                // 步骤3：向地址0x9002写1，等待100ms
                ProgressInfo = "开始擦除固件";
                await WriteRegister(0x9002, 1);
                await Task.Delay(200);

                // 步骤4：读取地址0x9003，判断值是否为1（擦除完成）
                ProgressInfo = "等待擦除完成";
                bool eraseCompleted = false;
                int maxRetries = 30;
                for (int i = 0; i < maxRetries; i++)
                {
                    int value = await ReadRegister(0x9003);
                    if (value == 1)
                    {
                        eraseCompleted = true;
                        break;
                    }
                    await Task.Delay(1000);
                }

                if (!eraseCompleted)
                {
                    throw new System.Exception("擦除固件失败");
                }

                // 步骤5：发送固件数据
                ProgressInfo = "发送固件数据";
                await SendFirmwareData();

                // 步骤6：读地址0x9002，判断9005值是否等于发送总数.
                ProgressInfo = "验证数据传输";
                bool verificationPassed = false;
                for (int i = 0; i < 3; i++)
                {
                    var values = await ReadMultipleRegisters(0x9002, 5);
                    if (values != null && values.Length >= 5 && values[2] == 1 && values[4] == 1 && values[3] == _packetCount) 
                    {
                        verificationPassed = true;
                        break;
                    }
                    await Task.Delay(1000);
                }

                if (!verificationPassed)
                {
                    throw new System.Exception("数据传输验证失败");
                }

                // 步骤7：向0x9002写3
                ProgressInfo = "完成升级";
                await WriteRegister(0x9002, 3);

                StatusMessage = "升级成功";
                ProgressValue = 100;
                ProgressInfo = "固件升级完成";
                _updateCompleted = true;
                MessageBox.Show("固件升级成功", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (System.Exception ex)
            {
                StatusMessage = "升级失败";
                //ProgressInfo = $"升级失败: {ex.Message}";
                MessageBox.Show($"升级失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// 写入单个寄存器
        /// </summary>
        /// <param name="address">寄存器地址</param>
        /// <param name="value">写入值</param>
        private async Task WriteRegister(int address, int value)
        {
            byte[] data = new byte[2];
            data[0] = (byte)(value >> 8);
            data[1] = (byte)(value & 0xFF);
            var payload = _otherConnectionService.BuildWritePayload(address, data);
            
            var result = _otherConnectionService.SendWriteMessage(payload);
            //if (!result.Succeeded)
            //{
            //    throw new System.Exception(result.Message);
            //}
            await Task.Delay(50);
        }

        /// <summary>
        /// 读取单个寄存器
        /// </summary>
        /// <param name="address">寄存器地址</param>
        /// <returns>寄存器值</returns>
        private async Task<int> ReadRegister(int address)
        {
            var payload = _otherConnectionService.BuildRealtimeReadPayload(address, 2);
            var response = await _otherConnectionService.SendRealtimeReadMessageAsync(payload);
            if (response.Length < 4 + 2)
            {
                throw new System.Exception("读取寄存器失败");
            }
            return (response[4] << 8) | response[5];
        }

        /// <summary>
        /// 写入多个寄存器
        /// </summary>
        /// <param name="address">起始寄存器地址</param>
        /// <param name="data">写入数据</param>
        private async Task WriteMultipleRegisters(int address, byte[] data)
        {
            var payload = _otherConnectionService.BuildWritePayload(address, data);
            Debug.WriteLine($"写入数据={ConvertToHexString(payload)}");
            var result = _otherConnectionService.SendWriteMessage(payload);
            if (!result.Succeeded)
            {
                throw new System.Exception(result.Message);
            }
            await Task.Delay(50);
        }

        /// <summary>
        /// 读取多个寄存器
        /// </summary>
        /// <param name="address">起始寄存器地址</param>
        /// <param name="count">寄存器数量</param>
        /// <returns>寄存器值数组</returns>
        private async Task<int[]> ReadMultipleRegisters(int address, int count)
        {
                var payload = _otherConnectionService.BuildRealtimeReadPayload(address, count * 2);
                var response = await _otherConnectionService.SendRealtimeReadMessageAsync(payload);
                if (response.Length < 4 + count * 2)
                {
                    Debug.WriteLine($"读取多个寄存器失败:{BitConverter.ToString( response)}");
                    return null;
                }
                int[] values = new int[count];
                for (int i = 0; i < count; i++)
                {
                    values[i] = (response[4 + i * 2] << 8) | response[4 + i * 2 + 1];
                }
                return values;
            }
           
        

        /// <summary>
        /// 计算CRC16校验值 (MODBUS标准)
        /// </summary>
        /// <param name="data">数据字节数组</param>
        /// <returns>CRC16校验值（高位在前）</returns>
        private byte[] CalculateCRC16(byte[] data)
        {
            int crc = 0xFFFF; // MODBUS标准初始值
            for (int i = 0; i < data.Length; i++)
            {
                crc ^= data[i];
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 0x0001) != 0)
                        crc = (crc >> 1) ^ 0xA001; // 反射多项式
                    else
                        crc >>= 1;
                }
            }
            return new byte[] { (byte)(crc >> 8), (byte)(crc & 0xFF) }; // 高位在前

        }
        private static string ConvertToHexString(byte[] payload) =>
       ConvertToHexString(payload, payload.Length);
        private static string ConvertToHexString(byte[] payload, int length) =>
       string.Join(" ", payload.Take(length).Select(value => value.ToString("X2")));

        /// <summary>
        /// 发送固件数据
        /// </summary>
        private async Task SendFirmwareData()
        {
            // 读取固件文件数据
            byte[] firmwareData = File.ReadAllBytes(FirmwareFilePath);
            int totalBytes = firmwareData.Length;
            int sentBytes = 0;
            int packetSize = 62; // 每次发送62字节
            _packetCount = 0;

            // 循环发送数据
            while (sentBytes < totalBytes)
            {
                // 计算本次发送的字节数
                int bytesToSend = Math.Min(packetSize, totalBytes - sentBytes);
                byte[] packet = new byte[packetSize];
                // 复制数据到发送缓冲区
                Array.Copy(firmwareData, sentBytes, packet, 0, bytesToSend);
                
                // 不足62字节的补FF
                for (int i = bytesToSend; i < packetSize; i++)
                {
                    packet[i] = 0xFF;
                }

                // 计算CRC16校验
                byte[] crc = CalculateCRC16(packet);
                // 构建完整的发送数据（62字节数据 + 2字节CRC）
                byte[] sendData = new byte[64];
                Array.Copy(packet, sendData, 62);
                Array.Copy(crc, 0, sendData, 62, 2);

                Debug.WriteLine($"发送的数据{ConvertToHexString(sendData)}");

                // 重发机制：最多尝试3次
                int retryCount = 0;
                const int maxRetries = 3;
                bool sendSuccess = false;

                while (retryCount < maxRetries && !sendSuccess)
                {
                    // 发送数据帧（ID201）
                    var result = _otherConnectionService.SendCanMessageWithId(0, sendData, 0x201);
                    if (!result.Succeeded)
                    {
                        retryCount++;
                        Debug.WriteLine($"发送数据失败: {result.Message}，正在重试 ({retryCount}/{maxRetries})");
                        await Task.Delay(50); // 等待50ms后重试
                        continue;
                    }

                    // 等待设备返回数据（采用应答模式，直接接收设备回复的报文）
                    bool ackReceived = false;

                    // 接收设备返回的报文
                    byte[] response = _otherConnectionService.ReceiveCanResponse(0);

                    Debug.WriteLine($"接收到的数据：{ConvertToHexString(response)}");

                    if (response.Length == 11)
                    {
                        // 检测9004和9006值是否为1
                        int reg9004 = (response[5] << 8) | response[6]; // 9002 + 2
                        int reg9006 = (response[9] << 8) | response[10]; // 9002 + 4
                        if (reg9004 == 1 && reg9006 == 1)
                        {
                            ackReceived = true;
                            sendSuccess = true;
                        }
                        else
                        {
                            retryCount++;
                            Debug.WriteLine($"设备确认失败，正在重试 ({retryCount}/{maxRetries})");
                            await Task.Delay(10); // 等待10ms后重试
                            continue;
                        }
                    }
                    else
                    {
                        retryCount++;
                        Debug.WriteLine($"接收响应长度不正确，正在重试 ({retryCount}/{maxRetries})");
                        await Task.Delay(10); // 等待10ms后重试
                        continue;
                    }

                    await Task.Delay(8);
                }

                if (!sendSuccess)
                {
                    throw new System.Exception($"数据发送失败，已尝试 {maxRetries} 次均未成功");
                }

                // 更新发送状态
                sentBytes += bytesToSend;
                _packetCount++;
                ProgressValue = (int)((double)sentBytes / totalBytes * 100);
                ProgressInfo = $"发送固件数据: {sentBytes}/{totalBytes} 字节";
            }
        }

        /// <summary>
        /// 取消升级命令
        /// </summary>
        [RelayCommand]
        private void Cancel()
        {
            StatusMessage = "已取消";
            ProgressInfo = "升级已取消";
        }

        /// <summary>
        /// 关闭窗口命令
        /// </summary>
        /// <summary>
        /// 窗口关闭时的处理逻辑
        /// </summary>
        public async Task OnWindowClosing()
        {
            // 检查升级是否完成，如果没有完成，执行步骤7确保设备正常关闭
            if (!_updateCompleted && StatusMessage == "升级中...")
            {
                try
                {
                    // 执行步骤7：向0x9002写3
                    await WriteRegister(0x9002, 3);
                    ProgressInfo = "已执行设备关闭操作";
                }
                catch (System.Exception ex)
                {
                    // 忽略错误，确保窗口能够关闭
                    Console.WriteLine($"执行设备关闭操作失败: {ex.Message}");
                }
            }
        }



        /// <summary>
        /// 获取设备信息命令
        /// </summary>
        [RelayCommand]
        private async Task GetDeviceInfo()
        {
            // 模拟获取设备信息
            var values = await ReadMultipleRegisters(0x9000, 2);
            if (values==null)
            {
                DeviceInfo = "获取错误";
                return;
            }
            DeviceInfo = $"版本号{values[0]}，固件号{values[1]}";
            StatusMessage = "已获取设备信息";
            ProgressInfo = "成功读取设备信息";

            await Task.Delay(1000);

        }
    }
}