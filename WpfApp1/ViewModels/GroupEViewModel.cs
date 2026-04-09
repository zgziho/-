using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.ViewModels
{
    public partial class GroupEViewModel:ObservableObject
    {
        [ObservableProperty]
        public string _rotateFreq="";
        [ObservableProperty]
        public string _currentRatio="";
        [ObservableProperty]
        public string _currentDigitalRef = "";
        [ObservableProperty]
        public string _currentPeak = "";
    }
}
