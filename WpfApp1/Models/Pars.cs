using CommunityToolkit.Mvvm.ComponentModel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Models
{
    /// <summary>
    /// 参数项模型：
    /// 与参数表Excel列一一对应，用于界面显示、读取和写入。
    /// </summary>
    public partial class Pars : ObservableObject
    {
        // 行唯一标识
        [ObservableProperty]
        private string? _id;
        // 参数地址（寄存器地址）
        [ObservableProperty]
        private string? _parsID;
        // 参数名称
        [ObservableProperty]
        private string? _parsNM;
        
        // 参数当前值
        [ObservableProperty]
        private string? _parsVA;
        // 单位
        [ObservableProperty]
        private string? _parsDW;
        // 类型
        [ObservableProperty]
        private string? _parsLX;
        // 属性（读/读写）
        [ObservableProperty]
        private string? _parsSX;
        // 范围
        [ObservableProperty]
        private string? _parsFW;
        // 系数方式
        [ObservableProperty]
        private string? _parsXSFS;
        // 系数
        [ObservableProperty]
        private string? _parsXS;
        // 系数位
        [ObservableProperty]
        private string? _parsXSW;
        // 类别
        [ObservableProperty]
        private string? _parsLB;
        // 颜色标识
        [ObservableProperty]
        private string? _parsYS;
        // 备注
        [ObservableProperty]
        private string? _parsBZ;

        
    }
}
