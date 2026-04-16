using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Windows;
using WpfApp1.Service;

namespace WpfApp1.ViewModels;


/// <summary>
/// 其他连接ViewModel类，用于管理设备连接和CAN通信相关的UI逻辑
/// 使用了CommunityToolkit.Mvvm库的ObservableObject和RelayCommand特性
/// </summary>
public partial class OtherConnectionViewModel : ObservableObject
{
    // 消息列表的最大行数，超过后会自动截断，防止内存占用过高
    private const int MaxMessageLines = 500;
    // 其他连接服务实例，用于实际执行与设备的通信操作
    private readonly OtherConnectionService _otherConnectionService;

    /// <summary>
    /// 连接状态显示文本，用于在UI上显示当前连接状态
    /// </summary>
    [ObservableProperty]
    private string _connectionStatus = "未连接";

    /// <summary>
    /// 消息列表内容，用于显示设备发送的消息
    /// </summary>
    [ObservableProperty]
    private string _listBoxContent = "";

    /// <summary>
    /// 打开设备按钮是否可用
    /// </summary>
    [ObservableProperty]
    private bool _isDeviceOpenEnabled = true;

    /// <summary>
    /// 初始化按钮是否可用
    /// </summary>
    [ObservableProperty]
    private bool _isInitializeEnabled;

    /// <summary>
    /// 启动按钮是否可用
    /// </summary>
    [ObservableProperty]
    private bool _isStartEnabled;

    /// <summary>
    /// 关闭设备按钮是否可用
    /// </summary>
    [ObservableProperty]
    private bool _isCloseEnabled;

    /// <summary>
    /// 是否使用自定义波特率
    /// </summary>
    [ObservableProperty]
    private bool _isCustomBaudrate;

    /// <summary>
    /// 自定义波特率文本框是否可见
    /// </summary>
    [ObservableProperty]
    private bool _isCustomBaudrateTetx;

    /// <summary>
    /// 自定义波特率值
    /// </summary>
    [ObservableProperty]
    private string _customBaudrate = "";

    /// <summary>
    /// 选中的仲裁波特率索引
    /// </summary>
    [ObservableProperty]
    private int _selectedArbitrationBaudrateIndex;

    /// <summary>
    /// 选中的数据波特率索引
    /// </summary>
    [ObservableProperty]
    private int _selectedDataBaudrateIndex;

    /// <summary>
    /// 是否为正常模式（0表示正常模式，1表示其他模式）
    /// </summary>
    [ObservableProperty]
    private int _isNormalMode;

    /// <summary>
    /// 是否为ISO标准（0表示ISO标准，1表示非ISO标准）
    /// </summary>
    [ObservableProperty]
    private int _isIsoStandard;

    /// <summary>
    /// 是否启用终端电阻
    /// </summary>
    [ObservableProperty]
    private bool _isTerminationEnable;

    /// <summary>
    /// 选中的过滤器模式索引，默认为2
    /// </summary>
    [ObservableProperty]
    private int _selectedFilterModeIndex = 2;

    /// <summary>
    /// 选中的通道索引
    /// </summary>
    [ObservableProperty]
    private int _selectedChannelIndex;

    /// <summary>
    /// 是否滚动到消息列表末尾，用于自动滚动到最新消息
    /// </summary>
    [ObservableProperty]
    private bool _scrollToEnd;

    /// <summary>
    /// 初始化其他连接ViewModel。
    /// </summary>
    /// <param name="otherConnectionService">其他连接服务实例，由依赖注入提供。</param>
    public OtherConnectionViewModel(OtherConnectionService otherConnectionService)
    {
        _otherConnectionService = otherConnectionService;
        // 订阅消息接收事件，当设备发送消息时会触发OnMessagesReceived方法
        //_otherConnectionService.MessagesReceived += OnMessagesReceived;
        // 设置当前活动通道索引
        _otherConnectionService.ActiveChannelIndex = SelectedChannelIndex;
        // 初始更新命令按钮状态，根据设备当前状态设置按钮可用性
        UpdateCommandStates();
    }
    
    /// <summary>
    /// 当IsCustomBaudrate属性改变时调用。
    /// 用于控制自定义波特率文本框的是否可见
    /// </summary>
    /// <param name="value">新的IsCustomBaudrate值。</param>
    partial void OnIsCustomBaudrateChanged(bool value)
    {
        // 当选择使用自定义波特率时，显示自定义波特率文本框
        IsCustomBaudrateTetx = value;
    }

    /// <summary>
    /// 当SelectedChannelIndex属性改变时调用。
    /// 用于更新服务中的活动通道索引
    /// </summary>
    /// <param name="value">新的通道索引值。</param>
    partial void OnSelectedChannelIndexChanged(int value)
    {
        // 将新的通道索引传递给服务
        _otherConnectionService.ActiveChannelIndex = value;
    }

    /// <summary>
    /// 打开设备命令，通过RelayCommand绑定到UI按钮
    /// </summary>
    [RelayCommand]
    private async Task OpenDevice()
    {
        // 调用服务的OpenDeviceAsync方法打开设备，返回操作结果
        var result = await _otherConnectionService.OpenDeviceAsync();
        // 处理操作结果，如果操作失败则返回
        if (!HandleResult(result))
        {
            return;
        }

        // 更新连接状态为已连接设备
        ConnectionStatus = "已连接设备";
        // 更新命令按钮状态，根据新的设备状态设置按钮可用性
        UpdateCommandStates();
    }

    /// <summary>
    /// 初始化CAN设备命令，通过RelayCommand绑定到UI按钮
    /// </summary>
    [RelayCommand]
    private void Initialize()
    {
        // 创建连接选项对象，包含所有必要的配置参数
        // 然后调用服务的Initialize方法初始化设备
        var result = _otherConnectionService.Initialize(new OtherConnectionOptions(
            IsCustomBaudrate,        // 是否使用自定义波特率
            CustomBaudrate,          // 自定义波特率值
            SelectedArbitrationBaudrateIndex,  // 仲裁波特率索引
            SelectedDataBaudrateIndex,        // 数据波特率索引
            IsNormalMode,            // 是否为正常模式
            IsIsoStandard,           // 是否为ISO标准
            IsTerminationEnable,     // 是否启用终端电阻
            SelectedFilterModeIndex));  // 过滤器模式索引

        // 处理操作结果，如果操作失败则返回
        if (!HandleResult(result))
        {
            return;
        }

        // 更新连接状态为CAN已初始化
        ConnectionStatus = "CAN已初始化";
        // 更新命令按钮状态，根据新的设备状态设置按钮可用性
        UpdateCommandStates();
    }

    /// <summary>
    /// 启动CAN设备命令，通过RelayCommand绑定到UI按钮
    /// </summary>
    [RelayCommand]
    private void Start()
    {
        // 调用服务的Start方法启动设备
        var result = _otherConnectionService.Start();
        // 处理操作结果，如果操作失败则返回
        if (!HandleResult(result))
        {
            return;
        }

        // 更新连接状态为CAN已启动
        ConnectionStatus = "CAN已启动";
        // 更新命令按钮状态，根据新的设备状态设置按钮可用性
        UpdateCommandStates();
    }

    /// <summary>
    /// 关闭设备命令，通过RelayCommand绑定到UI按钮
    /// </summary>
    [RelayCommand]
    private async Task CloseDevice()
    {
        // 调用服务的CloseDeviceAsync方法关闭设备
        var result = await _otherConnectionService.CloseDeviceAsync();
        // 处理操作结果，如果操作失败则返回
        if (!HandleResult(result))
        {
            return;
        }

        // 更新连接状态为未连接
        ConnectionStatus = "未连接";
        // 更新命令按钮状态，根据新的设备状态设置按钮可用性
        UpdateCommandStates();
    }

    /// <summary>
    /// 发送消息命令，通过RelayCommand绑定到UI按钮
    /// </summary>
    [RelayCommand]
    private void Sendmessage()
    {
        // 调用服务的SendDefaultMessage方法发送默认消息，指定通道索引
        var result = _otherConnectionService.SendDefaultMessage(SelectedChannelIndex);
        // 处理操作结果
        HandleResult(result);
    }


    /// <summary>
    /// 将消息添加到消息列表中
    /// 限制消息列表的最大行数，防止内存占用过高
    /// </summary>
    /// <param name="messages">要添加的消息列表</param>
    private void AppendMessages(IReadOnlyList<string> messages)
    {
        // 如果没有消息，直接返回
        if (messages.Count == 0)
        {
            return;
        }

        // 将现有消息分割成行，便于管理
        var existingLines = ListBoxContent
            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
            .ToList();

        // 添加新消息到现有消息列表
        existingLines.AddRange(messages);

        // 如果消息行数超过最大值，截取最新的MaxMessageLines行
        // 这样可以防止消息列表过长导致内存占用过高
        if (existingLines.Count > MaxMessageLines)
        {
            existingLines = existingLines
                .Skip(existingLines.Count - MaxMessageLines)  // 跳过旧消息
                .ToList();
        }

        // 更新消息列表内容，将所有消息行合并为一个字符串
        ListBoxContent = string.Join(Environment.NewLine, existingLines);
        // 切换ScrollToEnd值，触发UI滚动到底部，确保用户能看到最新消息
        ScrollToEnd = !ScrollToEnd;
    }

    /// <summary>
    /// 处理操作结果
    /// 统一处理各种操作的结果，包括成功和失败的情况
    /// </summary>
    /// <param name="result">操作结果对象</param>
    /// <returns>操作是否成功</returns>
    private bool HandleResult(OtherConnectionResult result)
    {
        // 如果操作成功，返回true
        if (result.Succeeded)
        {
            return true;
        }

        // 如果有详细信息，添加到消息列表，便于用户查看详细错误信息
        if (!string.IsNullOrWhiteSpace(result.Detail))
        {
            AppendMessages(new[] { result.Detail });
        }

        // 如果有错误消息，显示消息框，向用户提示错误
        if (!string.IsNullOrWhiteSpace(result.Message))
        {
            MessageBox.Show(result.Message);
        }

        // 更新命令按钮状态，确保按钮状态与设备当前状态一致
        UpdateCommandStates();
        // 操作失败，返回false
        return false;
    }

    /// <summary>
    /// 根据设备状态更新命令按钮的可用性
    /// 确保按钮的状态与设备当前状态匹配，提供良好的用户体验
    /// </summary>
    private void UpdateCommandStates()
    {
        // 如果设备未打开
        if (!_otherConnectionService.IsOpen)
        {
            // 只有打开设备按钮可用
            IsDeviceOpenEnabled = true;
            IsInitializeEnabled = false;
            IsStartEnabled = false;
            IsCloseEnabled = false;
            return;
        }

        // 如果设备已打开
        IsDeviceOpenEnabled = false;  // 禁用打开设备按钮，避免重复打开
        IsCloseEnabled = true;        // 启用关闭设备按钮
        // 只有当设备未初始化时，初始化按钮才可用
        IsInitializeEnabled = !_otherConnectionService.IsInitialized;
        // 只有当设备已初始化且未启动时，启动按钮才可用
        IsStartEnabled = _otherConnectionService.IsInitialized && !_otherConnectionService.IsStarted;
    }
}
