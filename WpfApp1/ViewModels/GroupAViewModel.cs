using CommunityToolkit.Mvvm.ComponentModel;

namespace WpfApp1.ViewModels
{
    /// <summary>
    /// 电流环参数分组ViewModel。
    /// </summary>
    public partial class GroupAViewModel : ObservableObject
    {
        // D轴比例系数
        [ObservableProperty]
        private string _dkp = "";

        // Q轴比例系数
        [ObservableProperty]
        private string _qkp = "";

        // 积分系数
        [ObservableProperty]
        private string _ki = "";

        // 转矩环频带
        [ObservableProperty]
        private string _torqueBand = "";
    }
}
