using CommunityToolkit.Mvvm.ComponentModel;

namespace WpfApp1.ViewModels
{
    /// <summary>
    /// 速度环参数分组ViewModel。
    /// </summary>
    public partial class GroupBViewModel : ObservableObject
    {
        // 速度环比例系数
        [ObservableProperty]
        private string _kp = "";

        // 速度环积分系数
        [ObservableProperty]
        private string _ki = "";

        // 前馈项
        [ObservableProperty]
        private string _tff = "";

        // 速度环频带
        [ObservableProperty]
        private string _speedBand = "";
    }
}
