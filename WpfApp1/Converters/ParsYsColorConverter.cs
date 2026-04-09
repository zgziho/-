using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace WpfApp1.Converters
{
    /// <summary>
    /// 将ParsYS字符串转换为对应的Brush
    /// </summary>
    public class ParsYSColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string parsYS && !string.IsNullOrWhiteSpace(parsYS))
            {
                // 根据ParsYS的值返回对应的颜色
                return GetColorFromParsYS(parsYS.Trim());
            }
            
            // 默认返回透明背景
            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 根据ParsYS字符串获取对应的颜色
        /// </summary>
        /// <param name="parsYS">ParsYS字符串</param>
        /// <returns>对应的Brush</returns>
        private Brush GetColorFromParsYS(string parsYS)
        {
            if (int.TryParse(parsYS, out int colorCode))
            {
              switch (colorCode)
                {
                    case 0: // 红色
                        return new SolidColorBrush(Color.FromRgb(255, 200, 200)); // 浅红色
                    
                    case 1: // 绿色
                        return new SolidColorBrush(Color.FromRgb(200, 255, 200)); // 浅绿色
                    
                    case 2: // 蓝色
                        return new SolidColorBrush(Color.FromRgb(200, 200, 255)); // 浅蓝色
                    
                    case 3: // 黄色
                        return new SolidColorBrush(Color.FromRgb(255, 255, 200)); // 浅黄色
                    
                    case 4: // 橙色
                        return new SolidColorBrush(Color.FromRgb(255, 220, 180)); // 浅橙色
                    
                    case 5: // 紫色
                        return new SolidColorBrush(Color.FromRgb(220, 200, 255)); // 浅紫色
                    
                    default: // 其他数字
                        return Brushes.Transparent; // 透明背景
                }
            }
            
            // 如果不是数字，返回透明背景
            return Brushes.Transparent;
        }
            
    }
}
    
