using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;

namespace WpfApp1.Converters
{
    /// <summary>
    /// 路径显示转换器：
    /// 将完整路径转换为文件名显示。
    /// </summary>
    public class PathConverter : IValueConverter
    {
        public static PathConverter GetFileName { get; } = new PathConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string path)
            {
                return Path.GetFileName(path);
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // 仅用于单向显示，不支持反向转换
            throw new NotImplementedException();
        }
    }
}
