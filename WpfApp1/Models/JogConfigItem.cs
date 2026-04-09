namespace WpfApp1.Models
{
    // 点动控制配置项：由 JogConfig.xlsx 每一行映射而来
    public class JogConfigItem
    {
        // 配置名称（使能/电流给定/速度给定/位置给定）
        public string Name { get; set; } = string.Empty;
        // 目标寄存器地址（支持 16 进制或 10 进制文本）
        public string AddressId { get; set; } = string.Empty;
        // 数据类型（uint16/int16/int32等）
        public string DataType { get; set; } = "uint16";
        // 开状态写入值（主要给"使能"使用）
        public string OnState { get; set; } = string.Empty;
        // 关状态写入值（主要给"使能"使用）
        public string OffState { get; set; } = string.Empty;
    }
}
