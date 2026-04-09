using CommunityToolkit.Mvvm.ComponentModel;

namespace WpfApp1.ViewModels
{
    /// <summary>
    /// 位置环参数分组ViewModel。
    /// </summary>
    public partial class GroupCViewModel : ObservableObject
    {
        // 位置环比例系数
        [ObservableProperty]
        private string _kp = "";

        // 速度前馈
        [ObservableProperty]
        private string _vff = "";

        // 当前位置显示
        [ObservableProperty]
        private string _currentPosition = "";

    }
}
