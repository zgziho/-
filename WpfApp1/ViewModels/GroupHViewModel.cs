using CommunityToolkit.Mvvm.ComponentModel;

namespace WpfApp1.ViewModels
{
    public partial class GroupHViewModel : ObservableObject
    {
        [ObservableProperty]
        public string _currentRef = "";

        [ObservableProperty]
        public string _signalFreq = "";
    }
}