using CommunityToolkit.Mvvm.ComponentModel;

namespace WpfApp1.ViewModels
{
    public partial class GroupGViewModel : ObservableObject
    {
        [ObservableProperty]
        public string _dutyCycle = "";

        [ObservableProperty]
        public string _rotateFreq = "";
    }
}