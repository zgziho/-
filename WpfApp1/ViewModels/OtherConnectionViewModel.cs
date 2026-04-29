using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.Generic;
using System.Windows;
using WpfApp1.Service;

namespace WpfApp1.ViewModels;


/// <summary>
/// 其他连接ViewModel类，用于管理设备连接和CAN通信相关的UI逻辑
/// 使用了CommunityToolkit.Mvvm库的ObservableObject和RelayCommand特性
/// </summary>
public partial class OtherConnectionViewModel : ObservableObject
{
    // 其他连接服务实例，用于实际执行与设备的通信操作
    private readonly OtherConnectionService _otherConnectionService;

    /// <summary>
    /// 连接状态显示文本，用于在UI上显示当前连接状态
    /// </summary>
    [ObservableProperty]
    private string _connectionStatus = "未连接";

    /// <summary>
    /// 合并后的操作按钮文本（打开设备/初始化CAN/启动CAN）
    /// </summary>
    [ObservableProperty]
    private string _operationButtonText = "打开设备";

    /// <summary>
    /// 操作按钮是否可用
    /// </summary>
    [ObservableProperty]
    private bool _isOperationEnabled = true;

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
    /// 统一操作命令，点击一次自动完成所有连接步骤（打开设备→初始化CAN→启动CAN）
    /// </summary>
    [RelayCommand]
    private async Task Operation()
    {
        IsOperationEnabled = false;
        
        // 步骤1：打开设备
        if (!_otherConnectionService.IsOpen)
        {
            var openResult = await _otherConnectionService.OpenDeviceAsync();
            if (!HandleResult(openResult))
            {
                IsOperationEnabled = true;
                return;
            }
            ConnectionStatus = "已连接设备";
        }

        // 步骤2：初始化CAN
        if (!_otherConnectionService.IsInitialized)
        {
            var initResult = _otherConnectionService.Initialize(new OtherConnectionOptions(
                IsCustomBaudrate,
                CustomBaudrate,
                SelectedArbitrationBaudrateIndex,
                SelectedDataBaudrateIndex,
                IsNormalMode,
                IsIsoStandard,
                IsTerminationEnable,
                SelectedFilterModeIndex));

            if (!HandleResult(initResult))
            {
                IsOperationEnabled = true;
                return;
            }
            ConnectionStatus = "CAN已初始化";
        }

        // 步骤3：启动CAN
        if (!_otherConnectionService.IsStarted)
        {
            var startResult = _otherConnectionService.Start();
            if (!HandleResult(startResult))
            {
                IsOperationEnabled = true;
                return;
            }
            ConnectionStatus = "CAN已启动";
        }

        UpdateCommandStates();
        IsOperationEnabled = true;
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

        // 如果有详细信息，显示消息框
        if (!string.IsNullOrWhiteSpace(result.Detail))
        {
            MessageBox.Show(result.Detail);
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
    /// 根据设备状态更新命令按钮的可用性和文本
    /// 确保按钮的状态与设备当前状态匹配，提供良好的用户体验
    /// </summary>
    private void UpdateCommandStates()
    {
        if (!_otherConnectionService.IsOpen)
        {
            // 设备未打开
            OperationButtonText = "打开设备";
            IsOperationEnabled = true;
            IsCloseEnabled = false;
        }
        else if (!_otherConnectionService.IsInitialized)
        {
            // 设备已打开但未初始化
            OperationButtonText = "初始化CAN";
            IsOperationEnabled = true;
            IsCloseEnabled = true;
        }
        else if (!_otherConnectionService.IsStarted)
        {
            // 设备已初始化但未启动
            OperationButtonText = "启动CAN";
            IsOperationEnabled = true;
            IsCloseEnabled = true;
        }
        else
        {
            // 设备已启动
            OperationButtonText = "运行中";
            IsOperationEnabled = false;
            IsCloseEnabled = true;
        }
    }
}
