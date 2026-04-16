using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using ZLGCAN;
using ZLGCANDemo;

namespace WpfApp1.Service;

/// <summary>
/// 其他连接服务类，用于管理与CAN总线设备的通信
/// 提供设备的打开、初始化、启动、关闭等操作，并处理CAN消息的发送和接收
/// </summary>
public sealed class OtherConnectionService
{
    private readonly ModbusService _modbusService;
    // 串行化 CAN 发送与写优先控制
    private readonly SemaphoreSlim _canAccessLock = new(1, 1);
    private int _pendingWriteRequests = 0;
    private int _writeCooldownMilliseconds = 10;
    private DateTime _readBlockedUntilUtc = DateTime.MinValue;

    

    /// <summary>
    /// 写后冷却时间（毫秒），用于短时抑制实时读，避免写后立即回读产生抖动
    /// </summary>
    public int WriteCooldownMilliseconds
    {
        get => _writeCooldownMilliseconds;
        set => _writeCooldownMilliseconds = value < 0 ? 0 : value;
    }
 
    /// <summary>
    /// 启动被动接收循环
    /// </summary>
    public void StartPassiveReceive()
    {
        if (_passiveReceiveCts != null)
            return;
        _passiveReceiveCts = new CancellationTokenSource();
        var token = _passiveReceiveCts.Token;

        // 生产者任务：负责接收数据并存入队列
        Task.Run(() =>
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    // 遍历所有通道尝试接收数据
                    for (int ch = 0; ch < ChannelCount-1; ch++)
                    {
                        if (token.IsCancellationRequested)
                            break;
                        if (!_channelStarted[ch])
                            continue;

                        try
                        {
                            
                              var  payloads = ReceiveAnyCanResponse(ch);
                            
                            
                            if (payloads != null && payloads.Count > 0)
                            {
                                foreach (var payload in payloads)
                                {
                                    if (payload != null && payload.Length > 0)
                                    {
                                        //StoreRawMessage($"通道{ch} 被动接收: {ConvertToHexString(payload)}");
                                        // 将数据入队，不直接触发事件，避免阻塞接收线程
                                        _frameQueue.Enqueue(payload);
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // 忽略单次接收错误，继续循环
                        }
                    }

                    
                }
            }
            catch
            {
                // 忽略循环异常
            }
        }, token);

        //消费者任务：负责从队列中取出数据并触发事件
       _consumerTask = Task.Run(() =>
       {
           try
           {
               while (!token.IsCancellationRequested)
               {
                   if (_frameQueue.TryDequeue(out var payload))
                   {
                       try
                       {
                           RealtimeFrameReceived?.Invoke(payload);
                       }
                       catch { }
                   }
               }
           }
           catch
           {

           }
       }, token);
    }

    /// <summary>
    /// 停止被动接收循环
    /// </summary>
    public void StopPassiveReceive()
    {
        try
        {
            _passiveReceiveCts?.Cancel();
            // 等待消费者任务完成
            _consumerTask?.Wait(1000);
            _passiveReceiveCts?.Dispose();
        }
        catch { }
        finally 
        { 
            _passiveReceiveCts = null;
            _consumerTask = null;
            // 清空队列
            while (_frameQueue.TryDequeue(out _)) { }
        }
    }
    
  
    /// <summary>
    /// CAN FD仲裁波特率数组
    /// </summary>
    private static readonly uint[] UsbCanFdAbit =
    {
        1000000,
        800000,
        500000,
        250000,
        125000,
        100000,
        50000,
        800000
    };

    /// <summary>
    /// CAN FD数据波特率数组
    /// </summary>
    private static readonly uint[] UsbCanFdDbit =
    {
        5000000,
        4000000,
        2000000,
        1000000,
        800000,
        500000,
        250000,
        125000,
        100000
    };

    /// <summary>
    /// 标准CAN帧最大数据长度
    /// </summary>
    private const int CanMaxDlen = 8;
    /// <summary>
    /// CAN FD帧最大数据长度
    /// </summary>
    private const int CanFdMaxDlen = 64;
    /// <summary>
    /// 扩展帧标志
    /// </summary>
    private const uint CanEffFlag = 0x80000000U;
    /// <summary>
    /// 远程帧标志
    /// </summary>
    private const uint CanRtrFlag = 0x40000000U;
    /// <summary>
    /// CAN ID掩码
    /// </summary>
    private const uint CanIdFlag = 0x1FFFFFFFU;
    /// <summary>
    /// 设备类型
    /// </summary>
    private const uint DeviceType = 41;
    /// <summary>
    /// 通道数量
    /// </summary>
    private const int ChannelCount = 2;

    private const byte ReadCommandCode = 0x40;
    private const byte Readresponse = 0x60;
    private const byte WriteCommandCode = 0x23;
    private const uint DefaultMessageId = 0x601;

    /// <summary>
    /// 通道句柄数组
    /// </summary>
    private readonly IntPtr[] _channelHandles = new IntPtr[ChannelCount];
    /// <summary>
    /// 通道启动状态数组
    /// </summary>
    private readonly bool[] _channelStarted = new bool[ChannelCount];
    /// <summary>
    /// 接收数据线程数组
    /// </summary>
    private readonly recvdatathread[] _recvDataThreads = new recvdatathread[ChannelCount];

    /// <summary>
    /// 设备句柄
    /// </summary>
    private IntPtr _deviceHandle;
    /// <summary>
    /// 设备是否打开
    /// </summary>
    private bool _isOpen;
    /// <summary>
    /// 设备是否初始化
    /// </summary>
    private bool _isInitialized;
    private readonly object _rawMessageLock = new();
    private readonly List<string> _rawReceivedMessages = new();
    private byte[] _currentReadPayload = { 0x40,0x10,0x00,0x01,0x00,0x00,0x00,0x00 };
    
    /// <summary>
    /// 事务进行中标识，用于防止消息竞争
    /// </summary>
    private volatile bool _transactionInProgress = false;
    
    /// <summary>
    /// CAN缓冲区访问锁，确保事务期间其他线程无法读取消息
    /// </summary>
    private readonly SemaphoreSlim _bufferAccessLock = new SemaphoreSlim(1, 1);
    // 被动接收循环控制
    private CancellationTokenSource? _passiveReceiveCts;
    // 用于解耦数据接收和事件处理的队列（生产者-消费者模式）
    private readonly ConcurrentQueue<byte[]> _frameQueue = new ConcurrentQueue<byte[]>();
    // 消费者任务
    private Task? _consumerTask;

    /// <summary>
    /// 当被动接收到任意 CAN 载荷时触发（原始载荷字节）
    /// </summary>
    public event Action<byte[]>? RealtimeFrameReceived;


    /// <summary>
    /// 获取设备是否打开
    /// </summary>
    public bool IsOpen => _isOpen;
    /// <summary>
    /// 获取设备是否初始化
    /// </summary>
    public bool IsInitialized => _isInitialized;
    /// <summary>
    /// 获取是否有通道已启动
    /// </summary>
    public bool IsStarted => _channelStarted.Any(started => started);
    
    /// <summary>
    /// 获取事务是否正在进行中
    /// </summary>
    public bool TransactionInProgress => _transactionInProgress;
    
    public int ActiveChannelIndex { get; set; }
    public string LastSentPayloadHex { get; private set; } = "00 11 22 33 44 55 66 77";

    /// <summary>
    /// 消息接收事件
    /// </summary>
    public event Action<IReadOnlyList<string>>? MessagesReceived;

    public OtherConnectionService(ModbusService modbusService)
    {
        _modbusService = modbusService;
    }


    //[DllImport("zlgcan.dll", EntryPoint = "ZCAN_ClearBuffer", CallingConvention = CallingConvention.Cdecl)]
    //public static extern uint ZCAN_ClearBuffer(object obj);
    /// <summary>
    /// 打开设备
    /// </summary>
    /// <returns>操作结果</returns>
    public async Task<OtherConnectionResult> OpenDeviceAsync()
    {
        // 如果设备已经打开，直接返回成功
        if (_isOpen)
        {
            return OtherConnectionResult.Success();
        }

        // 异步打开设备
        _deviceHandle = await Task.Run(() => Method.ZCAN_OpenDevice(DeviceType, 0, 0));
        // 检查设备是否打开成功
        if ((long)_deviceHandle == 0)
        {
            return OtherConnectionResult.Fail("打开设备失败,请检查设备类型和设备索引号是否正确");
        }

        // 更新设备状态
        _isOpen = true;
        _isInitialized = false;
        // 初始化通道句柄和启动状态
        Array.Fill(_channelHandles, IntPtr.Zero);
        Array.Fill(_channelStarted, false);

        return OtherConnectionResult.Success();
    }

    /// <summary>
    /// 初始化设备
    /// </summary>
    /// <param name="options">初始化选项</param>
    /// <returns>操作结果</returns>
    public OtherConnectionResult Initialize(OtherConnectionOptions options)
    {
        // 检查设备是否打开
        if (!_isOpen)
        {
            return OtherConnectionResult.Fail("设备还没打开");
        }

        // 遍历所有通道进行初始化
        for (var channelIndex = 0; channelIndex < ChannelCount; channelIndex++)
        {
            // 配置通道
            if (!ConfigureChannel(channelIndex, options, out var errorMessage))
            {
                return OtherConnectionResult.Fail(errorMessage);
            }

            // 创建通道配置
            var config = CreateChannelConfig(options.IsNormalMode);
            var configPointer = Marshal.AllocHGlobal(Marshal.SizeOf(config));

            try
            {
                // 将配置结构转换为指针并初始化CAN通道
                Marshal.StructureToPtr(config, configPointer, false);
                _channelHandles[channelIndex] = Method.ZCAN_InitCAN(_deviceHandle, (uint)channelIndex, configPointer);
            }
            finally
            {
                // 释放内存
                Marshal.FreeHGlobal(configPointer);
            }

            // 检查通道是否初始化成功
            if ((long)_channelHandles[channelIndex] == 0)
            {
                return OtherConnectionResult.Fail($"初始化CAN通道 {channelIndex} 失败");
            }

            // 如果需要启用终端电阻
            if (options.IsTerminationEnable && !SetResistanceEnable(channelIndex, true))
            {
                return OtherConnectionResult.Fail($"使能通道 {channelIndex} 终端电阻失败");
            }

            // 设置过滤器
            if (!SetFilter(channelIndex, options.SelectedFilterModeIndex))
            {
                return OtherConnectionResult.Fail($"设置通道 {channelIndex} 滤波失败");
            }
        }

        // 更新初始化状态
        _isInitialized = true;
        // 重置通道启动状态
        Array.Fill(_channelStarted, false);

        return OtherConnectionResult.Success();
    }

    /// <summary>
    /// 启动CAN设备
    /// </summary>
    /// <returns>操作结果</returns>
    public OtherConnectionResult Start()
    {
        // 检查设备是否打开
        if (!_isOpen)
        {
            return OtherConnectionResult.Fail("设备还没打开");
        }

        // 检查设备是否初始化
        if (!_isInitialized)
        {
            return OtherConnectionResult.Fail("请先初始化CAN");
        }

        // 遍历所有通道启动
        for (var channelIndex = 0; channelIndex < ChannelCount; channelIndex++)
        {
            // 启动CAN通道
            if (Method.ZCAN_StartCAN(_channelHandles[channelIndex]) != Define.STATUS_OK)
            {
                return OtherConnectionResult.Fail($"启动CAN通道 {channelIndex + 1} 失败");
            }

            // 更新通道启动状态
            _channelStarted[channelIndex] = true;


           // //如果接收线程不存在，创建并初始化
           // if (_recvDataThreads[channelIndex] is null)
           // {
           //     var receiveThread = new recvdatathread();
           //     var currentChannelIndex = channelIndex;
           //     // 订阅CAN数据接收事件
           //     receiveThread.RecvCANData += (data, len) => PublishCanMessages(data, len, currentChannelIndex);
           //     //// 订阅CAN FD数据接收事件
           //     receiveThread.RecvFDData += (data, len) => PublishCanFdMessages(data, len, currentChannelIndex);
           //     _recvDataThreads[channelIndex] = receiveThread;
           // }

           // //设置通道句柄并启动接收线程
           //_recvDataThreads[channelIndex].setChannelHandle(_channelHandles[channelIndex]);
           // _recvDataThreads[channelIndex].setStart(true);
        }

        return OtherConnectionResult.Success();
    }

    /// <summary>
    /// 关闭设备
    /// </summary>
    /// <returns>操作结果</returns>
    public async Task<OtherConnectionResult> CloseDeviceAsync()
    {
        // 停止接收线程
        StopReceiveThreads();

        // 如果设备已打开，关闭设备
        if (_isOpen)
        {
            await Task.Run(() => Method.ZCAN_CloseDevice(_deviceHandle));
        }

        // 重置设备状态
        _deviceHandle = IntPtr.Zero;
        _isOpen = false;
        _isInitialized = false;
        // 重置通道句柄和启动状态
        Array.Fill(_channelHandles, IntPtr.Zero);
        Array.Fill(_channelStarted, false);

        return OtherConnectionResult.Success();
    }

    /// <summary>
    /// 更新实时读取的消息负载数据
    /// </summary>
    /// <param name="payload">消息负载字节数组</param>
    public void UpdateRealtimeReadPayload(byte[] payload)
    {
        if (payload == null || payload.Length == 0)
        {
            return;
        }

        _currentReadPayload = payload.ToArray();
        LastSentPayloadHex = ConvertToHexString(_currentReadPayload);
    }

    public byte[] BuildRealtimeReadPayload(int startAddress, int count)
    {
        var payload = new byte[8];
        payload[0] = ReadCommandCode;
        payload[1] = (byte)((startAddress >> 8) & 0xFF);
        payload[2] = (byte)(startAddress & 0xFF);
        payload[3] = (byte)Math.Clamp(count, 0, byte.MaxValue);
        return payload;
    }
    /// <summary>
    /// 写入值
    /// </summary>
    /// <param name="startAddress"></param>
    /// <param name="data"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public byte[] BuildWritePayload(int startAddress, byte[] data)
    {
        data ??= Array.Empty<byte>();
        var payloadLength = Math.Max(CanMaxDlen, 4 + data.Length);
        if (payloadLength > CanFdMaxDlen)
        {
            throw new ArgumentOutOfRangeException(nameof(data), $"写入数据长度不能超过 {CanFdMaxDlen - 4} 字节。");
        }

        var payload = new byte[payloadLength];
        payload[0] = WriteCommandCode;
        payload[1] = (byte)((startAddress >> 8) & 0xFF);
        payload[2] = (byte)(startAddress & 0xFF);
        payload[3] = (byte)Math.Clamp(data.Length, 0, byte.MaxValue);
        Array.Copy(data, 0, payload, 4, data.Length);
        return payload;
    }

    /// <summary>
    /// 获取原始接收到的消息快照
    /// </summary>
    /// <returns>原始消息的只读列表</returns>
    public IReadOnlyList<string> GetRawReceivedMessagesSnapshot()
    {
        lock (_rawMessageLock)
        {
            return _rawReceivedMessages.ToList();
        }
    }

    /// <summary>
    /// 发送实时读取消息
    /// </summary>
    /// <param name="payload">消息负载字节数组</param>
    /// <returns>操作结果</returns>
    //public OtherConnectionResult SendRealtimeReadMessage(byte[] payload)
    //{
    //    UpdateRealtimeReadPayload(payload);

    //    // 发送前快速检查是否有待写请求或处于写后冷却期
    //    if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
    //    {
    //        return OtherConnectionResult.Fail("跳过读取：存在待写请求或写后冷却中");
    //    }

    //    // 获取串口锁（序列化所有CAN访问），并在获取后再次检查写请求状态
    //    _canAccessLock.Wait();
    //    try
    //    {
    //        if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
    //        {
    //            return OtherConnectionResult.Fail("跳过读取：存在待写请求或写后冷却中");
    //        }

    //        return SendMessage(ActiveChannelIndex, _currentReadPayload);
    //    }
    //    finally
    //    {
    //        try { _canAccessLock.Release(); } catch { }
    //    }
    //}

    /// <summary>
    /// 异步发送实时读取消息并等待响应
    /// </summary>
    /// <param name="payload">消息负载字节数组</param>
    /// <param name="timeoutMilliseconds">超时时间（毫秒）</param>
    /// <returns>响应数据</returns>
    public async Task<byte[]> SendRealtimeReadMessageAsync(byte[] payload)
    {
        UpdateRealtimeReadPayload(payload);

        // 发送前快速检查是否有待写请求或处于写后冷却期
        if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
        {
            //若有挂起写请求或写后冷却，则直接回退为空响应
            return Array.Empty<byte>();
        }

        // 执行CAN事务，发送请求并等待响应
        var response = await ExecuteCanTransactionAsync(ActiveChannelIndex, _currentReadPayload);
        
       

        return response;
    }

    public OtherConnectionResult SendWriteMessage(byte[] payload)
    {
        try
        {
            // 执行CAN事务，发送请求并等待响应
            var response = ExecuteCanTransactionAsync(ActiveChannelIndex, payload).Result;

            // 写成功后短时抑制读取
            ////_readBlockedUntilUtc = DateTime.UtcNow.AddMilliseconds(_writeCooldownMilliseconds);

            // 根据响应判断写入是否成功
            if (response.Length > 0 && ValidateCanResponseFormat(response))
            {
                return OtherConnectionResult.Success();
            }
            else
            {
                return OtherConnectionResult.Fail("写入失败：未收到有效响应");
            }
        }
        catch (Exception ex)
        {
            return OtherConnectionResult.Fail($"写入失败：{ex.Message}");
        }
    }

   

    /// <summary>
    /// 执行CAN事务，发送请求并等待响应
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="request">请求数据</param>
    /// <returns>响应数据</returns>
    private async Task<byte[]> ExecuteCanTransactionAsync(int channelIndex, byte[] request)
    {
        // 在异步路径中获取 CAN 访问锁，保证与写操作互斥
        await _canAccessLock.WaitAsync().ConfigureAwait(false);
        try
        {
            // 设置事务进行中标识，防止其他线程截获消息
            _transactionInProgress = true;
            
            //获取锁后再次判断写优先或写后冷却窗口
            if (Volatile.Read(ref _pendingWriteRequests) > 0 || DateTime.UtcNow < _readBlockedUntilUtc)
            {
                // 有写请求或处于冷却期，回退为空响应
                return Array.Empty<byte>();
            }
            
            // 获取缓冲区访问锁，确保事务期间其他线程无法读取CAN消息
            await _bufferAccessLock.WaitAsync().ConfigureAwait(false);

            try
            {
                // 发送请求（在持有锁期间发送并接收，保持事务性）
                var sendResult = SendMessage(channelIndex, request);
                if (!sendResult.Succeeded)
                {
                    throw new InvalidOperationException($"发送请求失败: {sendResult.Message}");
                }
                // 等待读取响应或超时
                var timeoutTask = Task.Delay(200);
                var readTask = Task.Run(() => ReceiveCanResponse(channelIndex));

                var finished = await Task.WhenAny(readTask, timeoutTask).ConfigureAwait(false);

                if (finished == readTask)
                {
                    try
                    {
                        var response = await readTask.ConfigureAwait(false);

                        return response ?? Array.Empty<byte>();
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("接收响应时发生错误", ex);
                    }
                }
            }
            finally
            {
                // 释放缓冲区访问锁
                _bufferAccessLock.Release();
            }

            // 超时，返回空数组
            return Array.Empty<byte>();
        }
        finally
        {
            // 重置事务进行中标识
            _transactionInProgress = false;
            try { _canAccessLock.Release(); } catch { }
        }
    }

    /// <summary>
    /// 接收CAN响应
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="timeoutMilliseconds">超时时间（毫秒）</param>
    /// <returns>响应数据</returns>
    private byte[] ReceiveCanResponse(int channelIndex)
    {
        
        var currentChannelHandle = _channelHandles[channelIndex];
        var startTime = DateTime.UtcNow;

        while (DateTime.UtcNow - startTime < TimeSpan.FromMilliseconds(200))
        {
            var datalen = Method.ZCAN_GetReceiveNum(currentChannelHandle, 0);
            var fddatalen = Method.ZCAN_GetReceiveNum(currentChannelHandle, 1);
            
            // 尝试接收标准CAN消息
            if (datalen > 0)
            {
                int size = Marshal.SizeOf(typeof(ZCAN_Receive_Data));
                var pointer = Marshal.AllocHGlobal((int)datalen * size);
                try
                {
                    uint received = Method.ZCAN_Receive(currentChannelHandle, pointer, (uint)datalen, 50);
                    if (received > 0)
                    {
                        for (uint i = 0; i < received; i++)
                        {
                            var data = Marshal.PtrToStructure<ZCAN_Receive_Data>(new IntPtr(pointer.ToInt64() + i * size));
                            var dlc = (int)data.frame.can_dlc;
                            if (dlc > 0 && data.frame.data != null)
                            {
                                var payload = new byte[dlc];
                                Array.Copy(data.frame.data, payload, dlc);
                                if (ValidateCanResponseFormat(payload))
                                {
                                    return payload;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录错误但继续尝试
                    Console.WriteLine($"接收CAN消息错误: {ex.Message}");
                }
                finally
                {
                    Marshal.FreeHGlobal(pointer);
                }
            }
            if (fddatalen > 0)
            {
                int size = Marshal.SizeOf(typeof(ZCAN_ReceiveFD_Data));
                var fdPointer = Marshal.AllocHGlobal((int)fddatalen * size);
                try
                {
                    uint received = Method.ZCAN_ReceiveFD(currentChannelHandle, fdPointer, (uint)fddatalen, 50);
                    if (received > 0)
                    {
                        for (uint i = 0; i < received; i++)
                        {
                            var data = Marshal.PtrToStructure<ZCAN_ReceiveFD_Data>(new IntPtr(fdPointer.ToInt64() + i * size));
                            string hexStr = data.frame.data[3].ToString("X2");
                            int dataLength = Convert.ToInt32(hexStr, 16);
                            var len = dataLength + 4;

                            if (len > 0 && data.frame.data != null)
                            {
                                var payload = new byte[len];
                                Array.Copy(data.frame.data, payload, len);
                                if (ValidateCanResponseFormat(payload))
                                {
                                    return payload;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 记录错误但继续尝试
                    Console.WriteLine($"接收CAN FD消息错误: {ex.Message}");
                }
                finally
                {
                    Marshal.FreeHGlobal(fdPointer);
                }
            }
        }
        return Array.Empty<byte>();

    }
    /// <summary>
    /// 与 ReceiveCanResponse 类似，但不验证格式，直接返回所有原始接收载荷（标准或 FD）
    /// </summary>
    /// <param name="channelIndex"></param>
    /// <returns>所有接收到的CAN消息列表</returns>
    private List<byte[]> ReceiveAnyCanResponse(int channelIndex)
    {
        // 如果有事务正在进行，返回空列表，避免截获指令响应
        if (_transactionInProgress)
        {
            return new List<byte[]>();
        }
        
        // 尝试获取缓冲区访问锁，如果获取失败（事务进行中），返回空列表
        if (!_bufferAccessLock.Wait(0))
        {
            return new List<byte[]>();
        }
        
        try
        {
            var currentChannelHandle = _channelHandles[channelIndex];
            var startTime = DateTime.UtcNow;
            var allMessages = new List<byte[]>();

            while (DateTime.UtcNow - startTime < TimeSpan.FromMilliseconds(50))
            {
            var datalen = Method.ZCAN_GetReceiveNum(currentChannelHandle, 0);
            var fddatalen = Method.ZCAN_GetReceiveNum(currentChannelHandle, 1);

            if (datalen > 0)
            {
                int size = Marshal.SizeOf(typeof(ZCAN_Receive_Data));
                var pointer = Marshal.AllocHGlobal((int)datalen * size);
                try
                {
                    uint received = Method.ZCAN_Receive(currentChannelHandle, pointer, (uint)datalen, 50);
                    if (received > 0)
                    {
                        for (uint i = 0; i < received; i++)
                        {
                            var data = Marshal.PtrToStructure<ZCAN_Receive_Data>(new IntPtr(pointer.ToInt64() + i * size));
                            var dlc = (int)data.frame.can_dlc;
                            if (dlc > 0 && data.frame.data != null)
                            {
                                var payload = new byte[dlc];
                                Array.Copy(data.frame.data, payload, dlc);
                                allMessages.Add(payload);
                            }
                        }
                    }
                }
                catch { }
                finally { Marshal.FreeHGlobal(pointer); }
            }

            if (fddatalen > 0)
            {
                int size = Marshal.SizeOf(typeof(ZCAN_ReceiveFD_Data));
                var fdPointer = Marshal.AllocHGlobal((int)fddatalen * size);
                try
                {
                    uint received = Method.ZCAN_ReceiveFD(currentChannelHandle, fdPointer, (uint)fddatalen, 50);
                   
                    if (received > 0)
                    {
                        for (uint i = 0; i < received; i++)
                        {
                            var data = Marshal.PtrToStructure<ZCAN_ReceiveFD_Data>(new IntPtr(fdPointer.ToInt64() + i * size));
                            var len = (int)data.frame.len;
                            if (len > 0 && data.frame.data != null)
                            {
                                var payload = new byte[len];
                                Array.Copy(data.frame.data, payload, len);
                                allMessages.Add(payload);
                            }
                        }
                    }
                }
                catch { }
                finally { Marshal.FreeHGlobal(fdPointer); }
            }
        }

        return allMessages;
        }
        finally
        {
            // 释放缓冲区访问锁
            _bufferAccessLock.Release();
        }
    }

    /// <summary>
    /// 验证CAN响应格式
    /// </summary>
    /// <param name="payload">响应数据</param>
    /// <returns>格式是否正确</returns>
    private bool ValidateCanResponseFormat(byte[] payload)
    {
        if (payload.Length < 4) 
        {
            return false;
        }

        // 检查命令码是否为0x60
        if (payload[0] != 0x60)
        {
            return false;
        }

        // 计算数据长度
        int dataLength = payload[3];
        
        // 验证总长度是否正确
        if (payload.Length != 4 + dataLength)
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// 发送默认消息
    /// </summary>
    /// <param name="selectedChannelIndex">选中的通道索引</param>
    /// <returns>操作结果</returns>
    public OtherConnectionResult SendDefaultMessage(int selectedChannelIndex)
    {
        ActiveChannelIndex = selectedChannelIndex;
        return SendMessage(selectedChannelIndex, _currentReadPayload);
    }

    /// <summary>
    /// 配置通道
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="options">配置选项</param>
    /// <param name="errorMessage">错误信息</param>
    /// <returns>配置是否成功</returns>
    private bool ConfigureChannel(int channelIndex, OtherConnectionOptions options, out string errorMessage)
    {
        // 设置CAN FD标准
        if (!SetCanFdStandard(channelIndex, options.IsIsoStandard))
        {
            errorMessage = "设置CANFD标准失败";
            return false;
        }

        // 根据是否使用自定义波特率进行配置
        if (options.IsCustomBaudrate)
        {
            if (!SetCustomBaudrate(channelIndex, options.CustomBaudrate))
            {
                errorMessage = "设置自定义波特率失败";
                return false;
            }
        }
        else
        {
            // 限制波特率索引范围
            var arbitrationIndex = Math.Clamp(options.SelectedArbitrationBaudrateIndex, 0, UsbCanFdAbit.Length - 1);
            var dataIndex = Math.Clamp(options.SelectedDataBaudrateIndex, 0, UsbCanFdDbit.Length - 1);

            // 设置CAN FD波特率
            if (!SetFdBaudrate(channelIndex, UsbCanFdAbit[arbitrationIndex], UsbCanFdDbit[dataIndex]))
            {
                errorMessage = "设置波特率失败";
                return false;
            }
        }

        errorMessage = string.Empty;
        return true;
    }

    /// <summary>
    /// 创建通道配置
    /// </summary>
    /// <param name="mode">模式</param>
    /// <returns>通道配置</returns>
    private static ZCAN_CHANNEL_INIT_CONFIG CreateChannelConfig(int mode)
    {
        var config = new ZCAN_CHANNEL_INIT_CONFIG();
        // 设置为CAN FD模式
        config.can_type = Define.TYPE_CANFD;
        // 设置模式
        config.canfd.mode = (byte)mode;
        return config;
    }

    /// <summary>
    /// 停止接收线程
    /// </summary>
    private void StopReceiveThreads()
    {
        // 遍历所有通道
        for (var channelIndex = 0; channelIndex < ChannelCount; channelIndex++)
        {
            // 如果接收线程存在且通道已启动，停止接收线程
            if (_recvDataThreads[channelIndex] is not null && _channelStarted[channelIndex])
            {
                _recvDataThreads[channelIndex].setStart(false);
            }

            // 重置通道启动状态
            _channelStarted[channelIndex] = false;
        }
    }

    /// <summary>
    /// 读取通道错误信息
    /// </summary>
    /// <param name="selectedChannelIndex">选中的通道索引</param>
    /// <returns>错误信息</returns>
    private string? ReadChannelError(int selectedChannelIndex)
    {
        var currentChannelHandle = _channelHandles[selectedChannelIndex];
        var errInfo = new ZCAN_CHANNEL_ERROR_INFO();
        var pointer = Marshal.AllocHGlobal(Marshal.SizeOf(errInfo));

        try
        {
            // 分配内存并读取错误信息
            Marshal.StructureToPtr(errInfo, pointer, false);
            if (Method.ZCAN_ReadChannelErrInfo(currentChannelHandle, pointer) != Define.STATUS_OK)
            {
                return "获取错误信息失败";
            }

            // 转换错误信息结构
            var errorInfo = Marshal.PtrToStructure<ZCAN_CHANNEL_ERROR_INFO>(pointer);
            return $"[通道{selectedChannelIndex}] 错误码：{errorInfo.error_code:D1}";
        }
        finally
        {
            // 释放内存
            Marshal.FreeHGlobal(pointer);
        }
    }


    /// <summary>
    /// 生成CAN ID
    /// </summary>
    /// <param name="id">基础ID</param>
    /// <param name="eff">是否为扩展帧</param>
    /// <param name="rtr">是否为远程帧</param>
    /// <param name="err">是否为错误帧</param>
    /// <returns>CAN ID</returns>
    private static uint MakeCanId(uint id, int eff, int rtr, int err)
    {
        var ueff = Convert.ToBoolean(eff) ? 1U : 0U;
        var urtr = Convert.ToBoolean(rtr) ? 1U : 0U;
        var uerr = Convert.ToBoolean(err) ? 1U : 0U;
        return id | ueff << 31 | urtr << 30 | uerr << 29;
    }

    private OtherConnectionResult SendMessage(int selectedChannelIndex, byte[] payload)
    {
        if (!IsValidChannelIndex(selectedChannelIndex))
        {
            return OtherConnectionResult.Fail("无效的通道索引");
        }

        if (!_channelStarted[selectedChannelIndex])
        {
            return OtherConnectionResult.Fail("请先启动CAN");
        }

        if (payload.Length == 0)
        {
            return OtherConnectionResult.Fail("发送数据为空");
        }

        if (payload.Length > CanFdMaxDlen)
        {
            return OtherConnectionResult.Fail("发送数据长度超过CANFD帧限制");
        }

        if (payload.Length <= CanMaxDlen)
        {
            return SendCanMessage(selectedChannelIndex, payload);
        }

        return SendCanFdMessage(selectedChannelIndex, payload);
    }

    private OtherConnectionResult SendCanMessage(int selectedChannelIndex, byte[] payload)
    {
        const int frameTypeIndex = 0;
        const int sendTypeIndex = 0;

        var currentChannelHandle = _channelHandles[selectedChannelIndex];
        uint result;
        var canData = new ZCAN_Transmit_Data
        { 
            frame =
            {
                can_id = MakeCanId(DefaultMessageId, frameTypeIndex, 0, 0),
                data = new byte[CanMaxDlen]
            },
            transmit_type = (uint)sendTypeIndex
        };

        Array.Copy(payload, canData.frame.data, payload.Length);
        canData.frame.can_dlc = (byte)payload.Length;

        var pointer = Marshal.AllocHGlobal(Marshal.SizeOf(canData));
        try
        {
            Marshal.StructureToPtr(canData, pointer, false);
            result = Method.ZCAN_Transmit(currentChannelHandle, pointer, 1);
        }
        finally
        {
            Marshal.FreeHGlobal(pointer);
        }

        if (result == 1)
        {
            LastSentPayloadHex = ConvertToHexString(payload);
            return OtherConnectionResult.Success();
        }

        var errorMessage = ReadChannelError(selectedChannelIndex) ?? "发送数据失败";
        return OtherConnectionResult.Fail("发送数据失败", errorMessage);
    }

    private OtherConnectionResult SendCanFdMessage(int selectedChannelIndex, byte[] payload)
    {
        const int frameTypeIndex = 0;
        const int sendTypeIndex = 0;

        var currentChannelHandle = _channelHandles[selectedChannelIndex];
        uint result;
        var canFdData = new ZCAN_TransmitFD_Data
        {
            frame =
            {
                can_id = MakeCanId(DefaultMessageId, frameTypeIndex, 0, 0),
                data = new byte[CanFdMaxDlen]
            },
            transmit_type = (uint)sendTypeIndex
        };

        Array.Copy(payload, canFdData.frame.data, payload.Length);
        canFdData.frame.len = (byte)payload.Length;

        var pointer = Marshal.AllocHGlobal(Marshal.SizeOf(canFdData));
        try
        {
            Marshal.StructureToPtr(canFdData, pointer, false);
            result = Method.ZCAN_TransmitFD(currentChannelHandle, pointer, 1);
        }
        finally
        {
            Marshal.FreeHGlobal(pointer);
        }

        if (result == 1)
        {
            LastSentPayloadHex = ConvertToHexString(payload);
            return OtherConnectionResult.Success();
        }

        var errorMessage = ReadChannelError(selectedChannelIndex) ?? "发送数据失败";
        return OtherConnectionResult.Fail("发送数据失败", errorMessage);
    }


    /// <summary>
    /// 将字节数组转换为十六进制字符串
    /// </summary>
    /// <param name="payload">字节数组</param>
    /// <returns>十六进制字符串</returns>
    private static string ConvertToHexString(byte[] payload) =>
        ConvertToHexString(payload, payload.Length);

    /// <summary>
    /// 将字节数组的指定长度转换为十六进制字符串
    /// </summary>
    /// <param name="payload">字节数组</param>
    /// <param name="length">转换长度</param>
    /// <returns>十六进制字符串</returns>
    private static string ConvertToHexString(byte[] payload, int length) =>
        string.Join(" ", payload.Take(length).Select(value => value.ToString("X2")));

    /// <summary>
    /// 设置过滤器
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="selectedFilterModeIndex">过滤器模式索引</param>
    /// <returns>设置是否成功</returns>
    private bool SetFilter(int channelIndex, int selectedFilterModeIndex)
    {
        // 清除过滤器
        if (!SetValue(channelIndex, "filter_clear", "0"))
        {
            return false;
        }

        var filterMode = selectedFilterModeIndex.ToString();
        // 如果是模式2，直接返回成功
        if (filterMode == "2")
        {
            return true;
        }

        // 设置过滤器模式
        if (!SetValue(channelIndex, "filter_mode", filterMode))
        {
            return false;
        }

        // 设置过滤器起始值
        if (!SetValue(channelIndex, "filter_start", "00000000"))
        {
            return false;
        }

        // 设置过滤器结束值
        if (!SetValue(channelIndex, "filter_end", "FFFFFFFF"))
        {
            return false;
        }

        // 设置过滤器确认
        return SetValue(channelIndex, "filter_ack", "0");
    }

    /// <summary>
    /// 设置终端电阻
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="isEnabled">是否启用</param>
    /// <returns>设置是否成功</returns>
    private bool SetResistanceEnable(int channelIndex, bool isEnabled) =>
        SetValue(channelIndex, "initenal_resistance", isEnabled ? "1" : "0");

    /// <summary>
    /// 设置CAN FD波特率
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="arbitrationBaud">仲裁波特率</param>
    /// <param name="dataBaud">数据波特率</param>
    /// <returns>设置是否成功</returns>
    private bool SetFdBaudrate(int channelIndex, uint arbitrationBaud, uint dataBaud) =>
        SetValue(channelIndex, "canfd_abit_baud_rate", arbitrationBaud.ToString()) &&
        SetValue(channelIndex, "canfd_dbit_baud_rate", dataBaud.ToString());

    /// <summary>
    /// 设置自定义波特率
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="baudrate">波特率</param>
    /// <returns>设置是否成功</returns>
    private bool SetCustomBaudrate(int channelIndex, string baudrate) =>
        !string.IsNullOrWhiteSpace(baudrate) && SetValue(channelIndex, "baud_rate_custom", baudrate);

    /// <summary>
    /// 设置CAN FD标准
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="canfdStandard">CAN FD标准</param>
    /// <returns>设置是否成功</returns>
    private bool SetCanFdStandard(int channelIndex, int canfdStandard) =>
        SetValue(channelIndex, "canfd_standard", canfdStandard.ToString());

    /// <summary>
    /// 设置设备值
    /// </summary>
    /// <param name="channelIndex">通道索引</param>
    /// <param name="suffix">参数后缀</param>
    /// <param name="value">值</param>
    /// <returns>设置是否成功</returns>
    private bool SetValue(int channelIndex, string suffix, string value)
    {
        var path = $"{channelIndex}/{suffix}";
        return Method.ZCAN_SetValue(_deviceHandle, path, Encoding.ASCII.GetBytes(value)) == 1;
    }

    /// <summary>
    /// 检查通道索引是否有效
    /// </summary>
    /// <param name="selectedChannelIndex">选中的通道索引</param>
    /// <returns>是否有效</returns>
    private static bool IsValidChannelIndex(int selectedChannelIndex) =>
        selectedChannelIndex >= 0 && selectedChannelIndex < ChannelCount;

    /// <summary>
    /// 发送实时读取报文
    /// </summary>
    public async Task<int[]?> SendRealtimeReadAndWaitForResponseAsync(int startAddress, int count)
    {
        try
        {
            var payload = BuildRealtimeReadPayload(startAddress, count*2);

            var response = await SendRealtimeReadMessageAsync(payload);

            // 解析响应数据
            if (response.Length > 0 && response != null)
            {
                if (response[0] == Readresponse)
                {
                    int parsedStart = (response[1] << 8) | response[2];
                    int parsedCount = response[3];
                    // 期望后续数据为 count 个 16 位寄存器（高字节在前）
                    if (parsedStart == startAddress && parsedCount == count * 2 && response.Length >= 4 + count * 2)
                    {
                        var values = new int[count];
                        for (int k = 0; k < count; k++)
                        {
                            int hi = response[4 + k * 2];
                            int lo = response[4 + k * 2 + 1];
                            values[k] = (hi << 8) | lo;
                        }
                        return values;
                    }
                }

            }

            return null;
        }
        catch (TimeoutException)
        {
            return null;
        }
        catch (Exception)
        {
            return null;
        }
    }
}

/// <summary>
/// 其他连接选项记录
/// </summary>
public sealed record OtherConnectionOptions(
    /// <summary>
    /// 是否使用自定义波特率
    /// </summary>
    bool IsCustomBaudrate,
    /// <summary>
    /// 自定义波特率值
    /// </summary>
    string CustomBaudrate,
    /// <summary>
    /// 选中的仲裁波特率索引
    /// </summary>
    int SelectedArbitrationBaudrateIndex,
    /// <summary>
    /// 选中的数据波特率索引
    /// </summary>
    int SelectedDataBaudrateIndex,
    /// <summary>
    /// 是否为正常模式
    /// </summary>
    int IsNormalMode,
    /// <summary>
    /// 是否为ISO标准
    /// </summary>
    int IsIsoStandard,
    /// <summary>
    /// 是否启用终端电阻
    /// </summary>
    bool IsTerminationEnable,
    /// <summary>
    /// 选中的过滤器模式索引
    /// </summary>
    int SelectedFilterModeIndex);

/// <summary>
/// 其他连接结果记录
/// </summary>
public sealed record OtherConnectionResult(
    /// <summary>
    /// 操作是否成功
    /// </summary>
    bool Succeeded,
    /// <summary>
    /// 消息
    /// </summary>
    string? Message = null,
    /// <summary>
    /// 详细信息
    /// </summary>
    string? Detail = null)
{
    /// <summary>
    /// 创建成功结果
    /// </summary>
    /// <returns>成功结果</returns>
    public static OtherConnectionResult Success() => new(true);

    /// <summary>
    /// 创建失败结果
    /// </summary>
    /// <param name="message">失败消息</param>
    /// <param name="detail">详细信息</param>
    /// <returns>失败结果</returns>
    public static OtherConnectionResult Fail(string message, string? detail = null) => new(false, message, detail);
}
