using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using WpfApp1.Views;

namespace WpfApp1.Service
{
    /// <summary>
    /// Modbus RTU 通信服务（纯 SerialPort 实现）：
    /// 提供连接、读写、CRC校验、读写仲裁、缓存回退与写优先控制。
    /// </summary>
    public partial class ModbusService : ObservableObject, IDisposable
    {
        /// <summary>
        /// 每次请求最大读取寄存器数量
        /// </summary>
        private const int MaxReadRegistersPerRequest = 120;
        /// <summary>
        /// 串口对象
        /// </summary>
        private SerialPort? _serialPort;
        /// <summary>
        /// 是否已释放
        /// </summary>
        private bool _isDisposed = false;
        /// <summary>
        /// 全局串口互斥：所有读写都必须串行执行，避免多窗口并发抢占串口
        /// </summary>
        private readonly SemaphoreSlim _serialAccessLock = new SemaphoreSlim(1, 1);
        /// <summary>
        /// 保持寄存器缓存：Key=寄存器地址，Value=寄存器值
        /// </summary>
        private readonly Dictionary<int, int> _holdingRegisterCache = new Dictionary<int, int>();
        /// <summary>
        /// 写操作完成后，短时间抑制实时读，避免设备尚未稳定时立即回读导致抖动
        /// </summary>
        private int _writeCooldownMilliseconds = 120;
        /// <summary>
        /// 读抑制截止时间（UTC），在该时间前优先走缓存不下发串口读
        /// </summary>
        private DateTime _readBlockedUntilUtc = DateTime.MinValue;
        /// <summary>
        /// 待写请求计数，用于实现“写优先”
        /// </summary>
        private int _pendingWriteRequests = 0;

        /// <summary>
        /// 可用端口列表
        /// </summary>
        public ObservableCollection<string> AvailablePorts { get; } = new ObservableCollection<string>();

        /// <summary>
        /// 连接状态
        /// </summary>
        [ObservableProperty]
        private bool _isConnected;

        /// <summary>
        /// 连接状态变更事件
        /// </summary>
        public event EventHandler<bool>? ConnectionStatusChanged;

        /// <summary>
        /// 写操作冷却时间（毫秒）
        /// </summary>
        public int WriteCooldownMilliseconds
        {
            get => _writeCooldownMilliseconds;
            // 防止外部传入负数，最小按 0ms 处理
            set => _writeCooldownMilliseconds = value < 0 ? 0 : value;
        }

        /// <summary>
        /// Modbus服务构造函数
        /// </summary>
        public ModbusService()
        {
            RefreshPorts();
        }

        /// <summary>
        /// 获取可用串口列表
        /// </summary>
        public ObservableCollection<string> RefreshPorts()
        {
            try
            {
                AvailablePorts.Clear();
                //获取串口信息并添加到列表
                foreach (string port in SerialPort.GetPortNames())
                {
                    AvailablePorts.Add(port);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取端口失败：{ex.Message}", "错误");
            }
            return AvailablePorts;
        }

         /// <summary>
         /// 连接串口
         /// </summary>
         /// <param name="portName">端口号</param>
         /// <param name="baudRate">波特率</param>
         /// <param name="parity">校验位</param>
         /// <param name="dataBits">数据位</param>
         /// <param name="stopBits">停止位</param>
         /// <param name="timeout">超时</param>
         /// <returns></returns>
        public async Task<bool> ConnectAsync(string portName, int baudRate, string parity, int dataBits, string stopBits, int timeout)
        {
            if (string.IsNullOrEmpty(portName) || _isDisposed) return false;
            if (SerialPort.GetPortNames().Length == 0)
            { 
                RefreshPorts();
                MessageBox.Show("串口已断开");
                return false; 
            }

            // 参数转换
            StopBits stopBit = GetStopBitsFromString(stopBits);
            System.IO.Ports.Parity portParity = GetParityFromString(parity);

            try
            {
                // 创建 SerialPort 对象 
                _serialPort = new SerialPort(portName, baudRate, portParity, dataBits, stopBit);
                _serialPort.ReadTimeout = timeout;
                _serialPort.WriteTimeout =timeout;

                // 打开串口 
                await Task.Run(() =>
                {
                    _serialPort.Open();
                });
                IsConnected = true;
                ConnectionStatusChanged?.Invoke(this, true);
                return true;
            }
            catch (Exception ex)
            {
                Cleanup();
                MessageBox.Show($"连接失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// 断开连接
        /// </summary>
        public void Disconnect()
        {

            Cleanup();
            IsConnected = false;
            ConnectionStatusChanged?.Invoke(this, false);
        }
        /// <summary>
        /// 批量读取
        /// </summary>
        /// <param name="slaveAddress">从机地址</param>
        /// <param name="startAddress">开始地址</param>
        /// <param name="numberOfPoints">读取数量</param>
        /// <returns></returns>
        public async Task<int[]?> ReadMultipleHoldingRegistersAsync(byte slaveAddress, string startAddress, int numberOfPoints)
        {
            if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
            {
                LogRead($"skip:not_connected slave={slaveAddress} start={startAddress} count={numberOfPoints}");
                return null;
            }
            if (numberOfPoints == 0)
                return Array.Empty<int>();
            if (!TryNormalizeHexAddress(startAddress, out var normalizedStart, out var parsedStart))
            {
                LogRead($"skip:invalid_address slave={slaveAddress} start={startAddress} count={numberOfPoints}");
                return null;
            }
            if (numberOfPoints > MaxReadRegistersPerRequest)
            {
                LogRead($"chunking:start slave={slaveAddress} start={normalizedStart} count={numberOfPoints}");
                var result = new int[numberOfPoints];
                int writeOffset = 0;
                int remaining = numberOfPoints;
                int currentStart = parsedStart;

                while (remaining > 0)
                {
                    int chunkCount = Math.Min(remaining, MaxReadRegistersPerRequest);
                    var chunk = await ReadMultipleHoldingRegistersAsync(slaveAddress, currentStart.ToString("X4"), chunkCount);
                    if (chunk == null || chunk.Length != chunkCount)
                    {
                        LogRead($"chunking:failed slave={slaveAddress} start={currentStart} count={chunkCount}");
                        return null;
                    }

                    Array.Copy(chunk, 0, result, writeOffset, chunkCount);
                    writeOffset += chunkCount;
                    remaining -= chunkCount;
                    currentStart += chunkCount;
                }

                LogRead($"chunking:success slave={slaveAddress} start={normalizedStart} count={numberOfPoints}");
                return result;
            }
            // 写优先：当有待写请求或处于写后冷却窗口时，优先返回缓存，避免与写抢占串口
            if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
            {
                LogRead($"fallback:precheck_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints} pendingWrite={Volatile.Read(ref _pendingWriteRequests)} blocked={(DateTime.UtcNow < _readBlockedUntilUtc)}");
                return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
            }

            try
            {
                // 串口访问全局串行化
                await _serialAccessLock.WaitAsync();
                if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                {
                    LogRead($"skip:not_connected_after_lock slave={slaveAddress} start={startAddress} count={numberOfPoints}");
                    return null;
                }
                // 获取锁后再次判断，避免锁等待期间写请求插入
                if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
                {
                    LogRead($"fallback:after_lock_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints} pendingWrite={Volatile.Read(ref _pendingWriteRequests)} blocked={(DateTime.UtcNow < _readBlockedUntilUtc)}");
                    return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
                }

                var timeoutTask = Task.Delay(_serialPort.ReadTimeout);

                var readTask = ReadHoldingRegistersCoreAsync(slaveAddress, normalizedStart, numberOfPoints);

                var completedTask = await Task.WhenAny(readTask, timeoutTask);
                if (completedTask == readTask)
                {
                    if (_serialPort == null || !_serialPort.IsOpen)
                        return null;

                    var values = await readTask;
                    // 真机读取成功后立即刷新缓存
                    UpdateCache(parsedStart, values);
                    LogRead($"success:device_read slave={slaveAddress} start={normalizedStart} count={numberOfPoints} values={values.Length}");
                    return values;
                }
                else
                {
                    // 读取超时时回退到缓存值，降低界面阻塞概率
                    LogRead($"fallback:timeout_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints}");
                    return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
                }
            }
            catch (Exception ex)
            {
                // 读取异常时同样回退到缓存
                LogRead($"fallback:exception_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints} ex={ex.GetType().Name}:{ex.Message}");
                return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
            }
            finally
            {
                if (_serialAccessLock.CurrentCount == 0)
                    _serialAccessLock.Release();
            }
        }
        /// <summary>
        /// 批量读取
        /// </summary>
        /// <param name="slaveAddress">从机地址</param>
        /// <param name="startAddress">开始地址</param>
        /// <param name="numberOfPoints">读取数量</param>
        /// <returns></returns>
        public async Task<int[]?> ReadMultipleHoldingRegistersAsync1(byte slaveAddress, string startAddress, int numberOfPoints)
        {
            if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
            {
                LogRead($"skip:not_connected slave={slaveAddress} start={startAddress} count={numberOfPoints}");
                return null;
            }
            if (numberOfPoints == 0)
                return Array.Empty<int>();
            if (!TryNormalizeHexAddress(startAddress, out var normalizedStart, out var parsedStart))
            {
                LogRead($"skip:invalid_address slave={slaveAddress} start={startAddress} count={numberOfPoints}");
                return null;
            }
            if (numberOfPoints > MaxReadRegistersPerRequest)
            {
                LogRead($"chunking:start slave={slaveAddress} start={normalizedStart} count={numberOfPoints}");
                var result = new int[numberOfPoints];
                int writeOffset = 0;
                int remaining = numberOfPoints;
                int currentStart = parsedStart;

                while (remaining > 0)
                {
                    int chunkCount = Math.Min(remaining, MaxReadRegistersPerRequest);
                    var chunk = await ReadMultipleHoldingRegistersAsync(slaveAddress, currentStart.ToString("X4"), chunkCount);
                    if (chunk == null || chunk.Length != chunkCount)
                    {
                        LogRead($"chunking:failed slave={slaveAddress} start={currentStart} count={chunkCount}");
                        return null;
                    }

                    Array.Copy(chunk, 0, result, writeOffset, chunkCount);
                    writeOffset += chunkCount;
                    remaining -= chunkCount;
                    currentStart += chunkCount;
                }

                LogRead($"chunking:success slave={slaveAddress} start={normalizedStart} count={numberOfPoints}");
                return result;
            }
            // 写优先：当有待写请求或处于写后冷却窗口时，优先返回缓存，避免与写抢占串口
            if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
            {
                LogRead($"fallback:precheck_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints} pendingWrite={Volatile.Read(ref _pendingWriteRequests)} blocked={(DateTime.UtcNow < _readBlockedUntilUtc)}");
                return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
            }

            try
            {
                // 串口访问全局串行化
                await _serialAccessLock.WaitAsync();
                if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                {
                    LogRead($"skip:not_connected_after_lock slave={slaveAddress} start={startAddress} count={numberOfPoints}");
                    return null;
                }
                // 获取锁后再次判断，避免锁等待期间写请求插入
                if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
                {
                    LogRead($"fallback:after_lock_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints} pendingWrite={Volatile.Read(ref _pendingWriteRequests)} blocked={(DateTime.UtcNow < _readBlockedUntilUtc)}");
                    return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
                }

                var timeoutTask = Task.Delay(_serialPort.ReadTimeout);

                var readTask = ReadHoldingRegistersCoreAsync();

                var completedTask = await Task.WhenAny(readTask, timeoutTask);
                if (completedTask == readTask)
                {
                    if (_serialPort == null || !_serialPort.IsOpen)
                        return null;

                    var values = await readTask;
                    return null;
                    // 真机读取成功后立即刷新缓存
                    //UpdateCache(parsedStart, readTask);
                    //LogRead($"success:device_read slave={slaveAddress} start={normalizedStart} count={numberOfPoints} values={values.Length}");
                    //return values;
                }
                   else
                {
                    // 读取超时时回退到缓存值，降低界面阻塞概率
                    LogRead($"fallback:timeout_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints}");
                    return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
                }
            }
            catch (Exception ex)
            {
                // 读取异常时同样回退到缓存
                LogRead($"fallback:exception_cache slave={slaveAddress} start={normalizedStart} count={numberOfPoints} ex={ex.GetType().Name}:{ex.Message}");
                return TryGetCachedHoldingRegisters(parsedStart, numberOfPoints, 2000);
            }
            finally
            {
                if (_serialAccessLock.CurrentCount == 0)
                    _serialAccessLock.Release();
            }
        }

        /// <summary>
        /// 写入单个保持寄存器
        /// </summary>
        public async Task<bool> WriteSingleRegisterAsync(byte slaveAddress, string registerAddress, string value)
        {
            if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                return false;

            // 先标记“有写请求”，让后续读逻辑主动让路
            Interlocked.Increment(ref _pendingWriteRequests);
            try
            {
                // 与读共用同一把锁，确保读写不会并发访问串口
                await _serialAccessLock.WaitAsync();
                if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                    return false;
                await WriteSingleRegisterCoreAsync(slaveAddress, registerAddress, value);
                // 写成功后同步更新缓存，避免参数表短时间显示旧值
                UpdateCache(registerAddress, value);
                // 启动写后冷却窗口：短时抑制实时读
                _readBlockedUntilUtc = DateTime.UtcNow.AddMilliseconds(_writeCooldownMilliseconds);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"写入寄存器失败：{ex.Message}", "错误");
                return false;
            }
            finally
            {
                if (_serialAccessLock.CurrentCount == 0)
                    _serialAccessLock.Release();
                // 写请求完成（成功或失败都要减计数）
                Interlocked.Decrement(ref _pendingWriteRequests);
            }
        }

        /// <summary>
        /// 批量写入多个保持寄存器（使用0x10功能码）
        /// </summary>
        /// <param name="slaveAddress">从站地址</param>
        /// <param name="startAddress">开始地址</param>
        /// <param name="values">要写入的值数组</param>
        /// <returns>写入是否成功</returns>
        public async Task<bool> WriteMultipleRegistersAsync(byte slaveAddress, string startAddress, ushort[] values)
        {
            if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                return false;

            // 先标记"有写请求"，让后续读逻辑主动让路
            Interlocked.Increment(ref _pendingWriteRequests);
            try
            {
                // 与读共用同一把锁，确保读写不会并发访问串口
                await _serialAccessLock.WaitAsync();
                if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                    return false;
                
                await WriteMultipleRegistersCoreAsync(slaveAddress, startAddress, values);
                
                // 写成功后同步更新缓存，避免参数表短时间显示旧值
                if (int.TryParse(startAddress, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int startAddr))
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        UpdateCache((startAddr + i).ToString("X4"), values[i].ToString("X4"));
                    }
                }
                
                // 启动写后冷却窗口：短时抑制实时读
                _readBlockedUntilUtc = DateTime.UtcNow.AddMilliseconds(_writeCooldownMilliseconds);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"批量写入寄存器失败：{ex.Message}", "错误");
                return false;
            }
            finally
            {
                if (_serialAccessLock.CurrentCount == 0)
                    _serialAccessLock.Release();
                // 写请求完成（成功或失败都要减计数）
                Interlocked.Decrement(ref _pendingWriteRequests);
            }
        }

        /// <summary>
        /// 按连续地址范围读取缓存
        /// </summary>
        /// <param name="startAddress">开始地址</param>
        /// <param name="numberOfPoints">读取数量</param>
        /// <param name="maxAgeMilliseconds">时间（不再使用）</param>
        /// <returns></returns>
        public int[]? TryGetCachedHoldingRegisters(int startAddress, int numberOfPoints, int maxAgeMilliseconds = 1000)
        {
            if (numberOfPoints == 0)
                return Array.Empty<int>();
            var result = new int[numberOfPoints];
            lock (_holdingRegisterCache)
            {
                for (int i = 0; i < numberOfPoints; i++)
                {
                    var address = startAddress + i;
                    if (!_holdingRegisterCache.TryGetValue(address, out var cached))
                    {
                        LogRead($"cache_miss start={startAddress} addr={address} count={numberOfPoints}");
                        return null;
                    }
                    result[i] = cached;
                }
            }
            LogRead($"cache_hit start={startAddress} count={numberOfPoints}");
            return result;
        }

        /// <summary>
        /// 获取单个地址的缓存值
        /// </summary>
        /// <param name="address">寄存器地址</param>
        /// <returns>缓存的字符串值，如果缓存中不存在则返回null</returns>
        public string? GetCachedValue(int address)
        {
            lock (_holdingRegisterCache)
            {
                if (_holdingRegisterCache.TryGetValue(address, out var cachedValue))
                {
                    return cachedValue.ToString();
                }
                return null;
            }
        }

        private void LogRead(string message)
        {
            Debug.WriteLine($"[ModbusRead] {DateTime.Now:HH:mm:ss.fff} {message}");
        }

        /// <summary>
        /// 用批量读结果刷新缓存
        /// </summary>
        private void UpdateCache(int startAddress, int[] values)
        {
            if (values == null || values.Length == 0)
                return;
            lock (_holdingRegisterCache)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    var address = startAddress + i;
                    _holdingRegisterCache[address] = values[i];
                }
            }
        }

        /// <summary>
        /// 允许外部（例如 CAN 接收）将批量读结果写入缓存的公开方法。
        /// 该方法只是对私有 UpdateCache 的薄封装。
        /// </summary>
        public void UpdateCacheFromExternal(int startAddress, int[] values)
        {
            UpdateCache(startAddress, values);
        }

        /// <summary>
        /// 用单点写结果刷新缓存
        /// </summary>
        private void UpdateCache(string addressHex, string valueHex)
        {
            if (!TryNormalizeHexAddress(addressHex, out _, out var address))
                return;
            if (!TryNormalizeHexAddress(valueHex, out _, out var value))
                return;
            lock (_holdingRegisterCache)
            {
                _holdingRegisterCache[address] = value;
            }
        }

        /// <summary>
        /// 清理资源
        /// </summary>
        private void Cleanup()
        {
            if (_serialPort != null)
            {
                try
                {
                    if (_serialPort.IsOpen)
                        _serialPort.Close();
                    _serialPort.Dispose();
                }
                catch { }
                _serialPort = null;
            }
            
            IsConnected = false;
            lock (_holdingRegisterCache)
            {
                _holdingRegisterCache.Clear();
            }
            _readBlockedUntilUtc = DateTime.MinValue;
            Interlocked.Exchange(ref _pendingWriteRequests, 0);
        }
        /// <summary>
        /// 自动检测波特率
        /// </summary>
        /// <param name="portName">端口号</param>
        /// <param name="baudRate">波特率</param>
        /// <param name="parity">校验码</param>
        /// <param name="dataBits">数据位</param>
        /// <param name="stopBits">停止位</param>
        /// <param name="timeout">超时时间</param>
        /// <param name="slaveAddress">从机地址</param>
        /// <param name="startAddress">起始地址</param>
        /// <param name="numberOfPoints">读取数量</param>
        /// <returns></returns>
        public async Task<bool> ProbeDeviceAsync(string portName, int baudRate, string parity, int dataBits, string stopBits, int timeout, byte slaveAddress, string startAddress, int numberOfPoints)
        {
            var stopBit = GetStopBitsFromString(stopBits);
            var portParity = GetParityFromString(parity);
            using var serialPort = new SerialPort(portName, baudRate, portParity, dataBits, stopBit)
            {
                ReadTimeout = timeout,
                WriteTimeout = timeout
            };

            try
            {
                await Task.Run(() => serialPort.Open());
                var values = await ReadHoldingRegistersCoreAsync(serialPort, slaveAddress, startAddress, numberOfPoints);
                return values.Length == numberOfPoints;
            }
            catch
            {
                return false;
            }
        }

      

        public async Task<byte[]> ReadHoldingRegistersCoreAsync()
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                throw new InvalidOperationException("串口未连接");
            return await ReadHoldingRegistersCoreAsync1();

        }
        private async Task<byte[]> ReadHoldingRegistersCoreAsync1( )
        {
            try
            {
                var response = await ExecuteRtuTransactionAsync();
                return response;


            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException($"读取保持寄存器超时: {ex.Message}", ex);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"读取保持寄存器协议错误: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"读取保持寄存器失败: {ex.Message}", ex);
            }


        }

        private async Task<byte[]> ExecuteRtuTransactionAsync()
        {
            try
            {
                return BuildChannelReadRequest(0x01);

            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException($"读取保持寄存器超时: {ex.Message}", ex);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"读取保持寄存器协议错误: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"读取保持寄存器失败: {ex.Message}", ex);
            }
        }


        /// <summary>
        /// 检测中途串口连接是否中断
        /// </summary>
        /// <param name="slaveAddress"></param>
        /// <param name="startAddress"></param>
        /// <param name="numberOfPoints"></param>
        /// <returns>处理好的数据</returns>
        /// <exception cref="InvalidOperationException"></exception>
        private async Task<int[]> ReadHoldingRegistersCoreAsync(byte slaveAddress, string startAddress, int numberOfPoints)
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                throw new InvalidOperationException("串口未连接");
            return await ReadHoldingRegistersCoreAsync(_serialPort, slaveAddress, startAddress, numberOfPoints);
        }
        /// <summary>
        /// 发送读取指令具体实现方法
        /// </summary>
        /// <param name="serialPort">串口对象</param>
        /// <param name="slaveAddress">从机地址</param>
        /// <param name="startAddress">起始地址</param>
        /// <param name="numberOfPoints">读取数量</param>
        /// <returns>高低位处理好的数据</returns>
        /// <exception cref="TimeoutException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        private async Task<int[]> ReadHoldingRegistersCoreAsync(SerialPort serialPort, byte slaveAddress, string startAddress, int numberOfPoints)
        {
            try
            {
                var request = BuildReadHoldingRegistersRequest(slaveAddress, startAddress, numberOfPoints);

                var response = await ExecuteRtuTransactionAsync(serialPort, request, 0x03);

                if (response.Length < 5)
                    throw new TimeoutException("读取响应长度不足。");

                if (response[1] == (0x03 | 0x80))
                    throw new InvalidOperationException($"设备返回异常码：0x{response[2]:X2}");

                if (response[2] != numberOfPoints * 2)
                    throw new InvalidOperationException("读取响应字节数与请求数量不一致。");

                var result = new int[numberOfPoints];
                for (int i = 0; i < numberOfPoints; i++)
                {
                    var high = response[3 + i * 2];
                    var low = response[4 + i * 2];
                    result[i] = (high << 8) | low;
                }
                return result;
            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException($"读取保持寄存器超时: {ex.Message}", ex);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"读取保持寄存器协议错误: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"读取保持寄存器失败: {ex.Message}", ex);
            }
        }
        /// <summary>
        /// 写单个寄存器06
        /// </summary>
        /// <param name="slaveAddress"></param>
        /// <param name="registerAddress"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <exception cref="TimeoutException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        private async Task WriteSingleRegisterCoreAsync(byte slaveAddress, string registerAddress, string value)
        {
            try
            {
                if (_serialPort == null || !_serialPort.IsOpen)
                    throw new InvalidOperationException("串口未连接");

                var request = BuildWriteSingleRegisterRequest(slaveAddress, registerAddress, value);

                var response = await ExecuteRtuTransactionAsync(_serialPort, request, 0x06);
                if (response.Length < 5)
                    throw new TimeoutException("写入响应长度不足。");
                if (response[1] == (0x06 | 0x80))
                    throw new InvalidOperationException($"设备返回异常码：0x{response[1]:X2}");
                if (response.Length != 8)
                    throw new InvalidOperationException("写入响应长度不正确。");
                for (int i = 0; i < 6; i++)
                {
                    if (response[i] != request[i])
                        throw new InvalidOperationException("写入响应与请求不一致。");
                }
            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException($"写入单个寄存器超时: {ex.Message}", ex);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"写入单个寄存器协议错误: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"写入单个寄存器失败: {ex.Message}", ex);
            }
        }
        /// <summary>
        /// 写多个寄存器10
        /// </summary>
        /// <param name="slaveAddress"></param>
        /// <param name="startAddress"></param>
        /// <param name="values"></param>
        /// <returns></returns>
        /// <exception cref="TimeoutException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        private async Task WriteMultipleRegistersCoreAsync(byte slaveAddress, string startAddress, ushort[] values)
        {
            try
            {
                if (_serialPort == null || !_serialPort.IsOpen)
                    throw new InvalidOperationException("串口未连接");

                var request = BuildWriteMultipleRegistersRequest(slaveAddress, startAddress, values);

                var response = await ExecuteRtuTransactionAsync(_serialPort, request, 0x10);
                if (response.Length < 5)
                    throw new TimeoutException("批量写入响应长度不足。");
                if (response[1] == (0x10 | 0x80))
                    throw new InvalidOperationException($"设备返回异常码：0x{response[1]:X2}");
                if (response.Length != 8)
                    throw new InvalidOperationException("批量写入响应长度不正确。");
                for (int i = 0; i < 6; i++)
                {
                    if (response[i] != request[i])
                        throw new InvalidOperationException("批量写入响应与请求不一致。");
                }
            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException($"批量写入寄存器超时: {ex.Message}", ex);
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"批量写入寄存器协议错误: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"批量写入寄存器失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 将数据转化为报文0X03
        /// </summary>
        /// <param name="slaveAddress">从机地址</param>
        /// <param name="startAddress">起始地址</param>
        /// <param name="numberOfPoints">数量</param>
        /// <returns>返回完整的报文</returns>
        private static byte[] BuildReadHoldingRegistersRequest(byte slaveAddress, string startAddress, int numberOfPoints)
        {
            byte[] byt = new byte[2];
            byt[0] = Convert.ToByte(startAddress.Substring(0, 2),16);
            byt[1] = Convert.ToByte(startAddress.Substring(2, 2), 16);
            
            var frame = new byte[8];
            frame[0] = slaveAddress;
            frame[1] = 0x03;
            frame[2] =byt[0];
            frame[3] = byt[1];
            frame[4] = (byte)((numberOfPoints >> 8) & 0xFF);
            frame[5] = (byte)(numberOfPoints & 0xFF);
            AppendCrc(frame);
            return frame;
        }

        /// <summary>
        /// 构建通道读取请求
        /// </summary>
        /// <param name="slaveAddress">从站地址</param>
        /// <returns>通道读取请求字节数组</returns>
        private static byte[] BuildChannelReadRequest(byte slaveAddress)
        {
            var frame = new byte[6];
            frame[0] = slaveAddress;
            frame[1] = 0x33;
            frame[2] = 0x00;
            frame[3] = 0x00;
            AppendCrc(frame);
            return frame;
        }

        /// <summary>
        /// 尝试标准化十六进制地址
        /// </summary>
        /// <param name="text">地址文本</param>
        /// <param name="normalized">标准化后的地址字符串</param>
        /// <param name="address">地址数值</param>
        /// <returns>是否标准化成功</returns>
        private static bool TryNormalizeHexAddress(string text, out string normalized, out int address)
        {
            normalized = string.Empty;
            address = 0;
            text = (text ?? string.Empty).Trim();
            if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                text = text[2..];
            if (!int.TryParse(text, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out address))
                return false;
            if (address < 0 || address > 0xFFFF)
                return false;
            normalized = address.ToString("X4");
            return true;
        }
        /// <summary>
        /// 将数据转化为报文0X06(传入值十进制)
        /// </summary>
        /// <param name="slaveAddress">从机地址</param>
        /// <param name="registerAddress">起始地址</param>
        /// <param name="value">值</param>
        /// <returns>返回完整的报文</returns>
        private static byte[] BuildWriteSingleRegisterRequest(byte slaveAddress, string registerAddress, string value)
        {
            byte[] byt = new byte[2];
            byt[0] = Convert.ToByte(registerAddress.Substring(0, 2), 16);
            byt[1] = Convert.ToByte(registerAddress.Substring(2, 2), 16);
            byte[] byt1 = new byte[2];
            byt1[0] = Convert.ToByte(value.Substring(0, 2), 16);
            byt1[1] = Convert.ToByte(value.Substring(2, 2), 16);
            var frame = new byte[8];
            frame[0] = slaveAddress;
            frame[1] = 0x06;
            frame[2] = byt[0];
            frame[3] = byt[1];
            frame[4] = byt1[0];
            frame[5] = byt1[1];
            AppendCrc(frame);
            return frame;
        }

        /// <summary>
        /// 将数据转化为报文0X10（批量写入多个寄存器）
        /// </summary>
        /// <param name="slaveAddress">从机地址</param>
        /// <param name="startAddress">起始地址</param>
        /// <param name="values">要写入的值数组</param>
        /// <returns>返回完整的报文</returns>
        private static byte[] BuildWriteMultipleRegistersRequest(byte slaveAddress, string startAddress, ushort[] values)
        {
            if (values.Length == 0)
                throw new ArgumentException("值数组不能为空");

            // 解析起始地址
            byte[] addressBytes = new byte[2];
            addressBytes[0] = Convert.ToByte(startAddress.Substring(0, 2), 16);
            addressBytes[1] = Convert.ToByte(startAddress.Substring(2, 2), 16);

            // 寄存器数量
            ushort registerCount = (ushort)values.Length;
            byte[] countBytes = BitConverter.GetBytes(registerCount);
            if (BitConverter.IsLittleEndian)
                Array.Reverse(countBytes);

            // 字节数（每个寄存器2字节）
            byte byteCount = (byte)(values.Length * 2);

            // 构建值数据
            byte[] valueBytes = new byte[byteCount];
            for (int i = 0; i < values.Length; i++)
            {
                byte[] registerBytes = BitConverter.GetBytes(values[i]);
                if (BitConverter.IsLittleEndian)
                    Array.Reverse(registerBytes);
                Array.Copy(registerBytes, 0, valueBytes, i * 2, 2);
            }

            // 构建完整报文
            var frame = new byte[9 + byteCount]; // 地址(1) + 功能码(1) + 起始地址(2) + 寄存器数量(2) + 字节数(1) + 值数据(byteCount) + CRC(2)
            frame[0] = slaveAddress;
            frame[1] = 0x10; // 功能码：写多个寄存器
            frame[2] = addressBytes[0];
            frame[3] = addressBytes[1];
            frame[4] = countBytes[0];
            frame[5] = countBytes[1];
            frame[6] = byteCount;
            
            // 复制值数据
            Array.Copy(valueBytes, 0, frame, 7, byteCount);
            
            AppendCrc(frame);
            return frame;
        }

        /// <summary>
        /// 发送报文并接收
        /// </summary>
        /// <param name="serialPort">串口对象</param>
        /// <param name="request">需要发送的报文</param>
        /// <param name="functionCode">功能码</param>
        /// <returns>接收到的报文</returns>
        /// <exception cref="InvalidOperationException"></exception>
        private async Task<byte[]> ExecuteRtuTransactionAsync(SerialPort serialPort, byte[] request, byte functionCode)
        {
            return await Task.Run(() =>
            {
                lock (serialPort)
                {
                    try
                    {
                        //清空接收缓冲区
                        serialPort.DiscardInBuffer();
                        //写入，从0开始写完
                        serialPort.Write(request, 0, request.Length);
                        //获取报文头
                        var head = ReadExact(serialPort, 2);
                        if (head[0] != request[0])
                            throw new InvalidOperationException("响应从机地址不匹配。");
                        var responseFunction = head[1];
                        //如果接收到的功能码为错误码
                        if (responseFunction == (functionCode | 0x80))
                        {//继续读取全部，确保错误码也是完整的
                            var tail = ReadExact(serialPort, 4);
                            var frame = new byte[6];
                            frame[0] = head[0];
                            frame[1] = head[1];
                            Array.Copy(tail, 0, frame, 2, 4);
                            ValidateCrc(frame);
                            return frame;
                        }

                        if (responseFunction != functionCode)
                            throw new InvalidOperationException($"响应功能码不匹配：0x{responseFunction:X2}");

                        byte[] frameData;
                     
                        //功能码正确
                        if (functionCode == 0x03)
                        {
                            //后面还有多少数据
                            var byteCount = ReadExact(serialPort, 1)[0];
                            //数据加上crc校验的两个
                            var bodyAndCrc = ReadExact(serialPort, byteCount + 2);
                            frameData = new byte[3 + bodyAndCrc.Length];
                            frameData[0] = head[0];
                            frameData[1] = head[1];
                            frameData[2] = byteCount;
                            Array.Copy(bodyAndCrc, 0, frameData, 3, bodyAndCrc.Length);
                        }
                        else
                        {
                            var bodyAndCrc = ReadExact(serialPort, 6);
                            frameData = new byte[8];
                            frameData[0] = head[0];
                            frameData[1] = head[1];
                            Array.Copy(bodyAndCrc, 0, frameData, 2, 6);
                        }
                        
                        ValidateCrc(frameData);
                        return frameData;
                    }
                    catch (TimeoutException ex)
                    {
                        throw new TimeoutException($"Modbus通信超时: {ex.Message}", ex);
                    }
                    catch (InvalidOperationException ex)
                    {
                        throw new InvalidOperationException($"Modbus协议错误: {ex.Message}", ex);
                    }
                    catch (IOException ex)
                    {
                        throw new InvalidOperationException($"串口通信错误: {ex.Message}", ex);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"Modbus通信失败: {ex.Message}", ex);
                    }
                }
            });
        }

        /// <summary>
        /// 发送报文并接收
        /// </summary>
        /// <param name="serialPort">串口对象</param>
        /// <param name="request">需要发送的报文</param>
        /// <param name="functionCode">功能码</param>
        /// <returns>接收到的报文</returns>
        /// <exception cref="InvalidOperationException"></exception>
        private async Task<byte[]> ExecuteRtuTransactionAsync1(SerialPort serialPort, byte[] request, byte functionCode)
        {
            return await Task.Run(() =>
            {
                lock (serialPort)
                {
                    try
                    {
                        //清空接收缓冲区
                        serialPort.DiscardInBuffer();
                        //写入，从0开始写完
                        serialPort.Write(request, 0, request.Length);
                        //获取报文头
                        var head = ReadExact(serialPort, 2);
                        if (head[0] != request[0])
                            throw new InvalidOperationException("响应从机地址不匹配。");
                        var responseFunction = head[1];

                        if (responseFunction != functionCode)
                            throw new InvalidOperationException($"响应功能码不匹配：0x{responseFunction:X2}");

                        byte[] frameData;
                        byte[] frameData1;
                        byte[] frameData2;

                      
                            //后面还有多少数据
                            var byteCount = ReadExact(serialPort, 1)[0];
                            //数据加上crc校验的两个
                            var bodyAndCrc = ReadExact(serialPort, byteCount);
                            frameData = new byte[3 + bodyAndCrc.Length];
                            frameData[0] = head[0];
                            frameData[1] = head[1];
                            frameData[2] = byteCount;
                            Array.Copy(bodyAndCrc, 0, frameData, 3, bodyAndCrc.Length);
                           
                            head = ReadExact(serialPort, 2);
                            
                                byteCount = ReadExact(serialPort, 1)[0];

                                bodyAndCrc = ReadExact(serialPort, byteCount + 2);
                                frameData1 = new byte[frameData.Length + bodyAndCrc.Length];
                                frameData1[0] = head[0];
                                frameData1[1] = head[1];
                                frameData1[2] = byteCount;
                                Array.Copy(bodyAndCrc, 0, frameData, 3, bodyAndCrc.Length);
                                frameData2 = new byte[frameData.Length + bodyAndCrc.Length+2];
                                Array.Copy(frameData, 0, frameData2, 0, frameData.Length);
                                Array.Copy(frameData1, 0, frameData2, frameData.Length, frameData1.Length);
                           
                        
                       

                        ValidateCrc(frameData2);
                        return frameData2;
                    }
                    catch (TimeoutException ex)
                    {
                        throw new TimeoutException($"Modbus通信超时: {ex.Message}", ex);
                    }
                    catch (InvalidOperationException ex)
                    {
                        throw new InvalidOperationException($"Modbus协议错误: {ex.Message}", ex);
                    }
                    catch (IOException ex)
                    {
                        throw new InvalidOperationException($"串口通信错误: {ex.Message}", ex);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"Modbus通信失败: {ex.Message}", ex);
                    }
                }
            });
        }
        /// <summary>
        /// 获取指定数量的报文
        /// </summary>
        /// <param name="serialPort">串口对象</param>
        /// <param name="length">获取多少个</param>
        /// <returns>报文</returns>
        /// <exception cref="TimeoutException"></exception>
        private static byte[] ReadExact(SerialPort serialPort, int length)
        {
            var buffer = new byte[length];
            var offset = 0;
            while (offset < length)
            {
                var read = serialPort.Read(buffer, offset, length - offset);
                if (read <= 0)
                    throw new TimeoutException("串口读取超时。");
                offset += read;
            }
            return buffer;
        }
        /// <summary>
        /// 将crc拼接到报文末端
        /// </summary>
        /// <param name="frame"></param>
        private static void AppendCrc(byte[] frame)
        {
            var crc = ComputeModbusCrc(frame, frame.Length - 2);
            frame[frame.Length - 2] = (byte)(crc & 0xFF);
            frame[frame.Length - 1] = (byte)(crc >> 8);
        }
        /// <summary>
        /// 对比回来的CRC校验是否正确
        /// </summary>
        /// <param name="frame"></param>
        /// <exception cref="InvalidOperationException"></exception>
        private static void ValidateCrc(byte[] frame)
        {
            var crc = ComputeModbusCrc(frame, frame.Length - 2);
            var low = (byte)(crc & 0xFF);
            var high = (byte)(crc >> 8);
            if (frame[^2] != low || frame[^1] != high)
                throw new InvalidOperationException($"CRC校验失败：期望 {low:X2} {high:X2}，实际 {frame[^2]:X2} {frame[^1]:X2}");
        }
        /// <summary>
        /// crc校验计算
        /// </summary>
        /// <param name="data"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        private static int ComputeModbusCrc(byte[] data, int length)
        {
            int crc = 0xFFFF;
            for (int i = 0; i < length; i++)
            {
                crc ^= data[i];
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 0x0001) != 0)
                        crc = (crc >> 1) ^ 0xA001;
                    else
                        crc >>= 1;
                }
            }
            return crc & 0xFFFF;
        }

        /// <summary>
        /// 停止位字符串转换为枚举
        /// </summary>
        private StopBits GetStopBitsFromString(string stopBit)
        {
            return stopBit switch
            {
                "1" => StopBits.One,
                "1.5" => StopBits.OnePointFive,
                "2" => StopBits.Two,
                _ => StopBits.One
            };
        }

        /// <summary>
        /// 校验位字符串转换为枚举
        /// </summary>
        private System.IO.Ports.Parity GetParityFromString(string parity)
        {
            return parity switch
            {
                "无" => System.IO.Ports.Parity.None,
                "奇校验" => System.IO.Ports.Parity.Odd,
                "偶校验" => System.IO.Ports.Parity.Even,
                _ => System.IO.Ports.Parity.None
            };
        }

        /// <summary>
        /// 发送自定义报文（功能码33）
        /// </summary>
        /// <param name="slaveAddress">从站地址</param>
        /// <param name="functionCode">功能码</param>
        /// <param name="customData">自定义数据</param>
        /// <returns>响应数据</returns>
        public async Task<byte[]> SendCustomMessageAsync(byte slaveAddress, byte functionCode, byte[] customData)
        {
            if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                return null;

            try
            {
                await _serialAccessLock.WaitAsync();
                if (_serialPort == null || !_serialPort.IsOpen || !IsConnected)
                    return null;

                // 构建自定义报文
                byte[] request = BuildCustomMessageRequest(slaveAddress, functionCode, customData);
                
                // 发送并接收响应
                var response = await ExecuteRtuTransactionAsync(_serialPort, request, functionCode);
                
                return response;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SendCustomMessageAsync] 发送自定义报文失败: {ex.Message}");
                return null;
            }
            finally
            {
                if (_serialAccessLock.CurrentCount == 0)
                    _serialAccessLock.Release();
            }
        }

        /// <summary>
        /// 构建自定义报文请求
        /// </summary>
        private byte[] BuildCustomMessageRequest(byte slaveAddress, byte functionCode, byte[] customData)
        {
            // 报文结构：从站地址 + 功能码 + 自定义数据 + CRC
            byte[] request = new byte[3 + customData.Length + 2]; // 头部3字节 + 数据 + CRC2字节
            
            request[0] = slaveAddress;          // 从站地址
            request[1] = functionCode;         // 功能码
            request[2] = (byte)customData.Length; // 数据长度
            
            // 复制自定义数据
            Array.Copy(customData, 0, request, 3, customData.Length);
            
            // 计算CRC
            ushort crc = (ushort)ComputeModbusCrc(request, request.Length - 2);
            request[request.Length - 2] = (byte)(crc & 0xFF);
            request[request.Length - 1] = (byte)(crc >> 8);
            
            return request;
        }

        public void Dispose()
        {
            if (!_isDisposed)
            {
                Cleanup();
                _isDisposed = true;
            }
        }
    }
}
