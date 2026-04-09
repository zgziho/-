using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WpfApp1.Models;

namespace WpfApp1.Service
{
    // 点动/使能配置服务：
    // 1) 启动时确保配置文件存在；
    // 2) 将 Excel 行映射为 JogConfigItem 集合供主窗体控制流程使用。
    public class JogConfigService
    {
        private readonly string _filePath;

        public JogConfigService(string? filePath = null)
        {
            // 默认放在程序运行目录，方便现场直接编辑
            _filePath = filePath ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "JogConfig.xlsx");
        }

        public IReadOnlyList<JogConfigItem> Load()
        {
            // 首次运行自动生成模板，避免“配置表还没创建”导致流程不可用
            EnsureTemplate();
            var list = new List<JogConfigItem>();
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed().Skip(1);
            foreach (var row in rows)
            {
                // 列映射：名称/地址id/开状态/关状态
                var item = new JogConfigItem
                {
                    Name = row.Cell(1).GetString().Trim(),
                    AddressId = row.Cell(2).GetString().Trim(),
                    DataType = row.Cell(3).GetString().Trim(),
                    OnState = row.Cell(4).GetString().Trim(),
                    OffState = row.Cell(5).GetString().Trim()
                };
                // 跳过空名称行，防止无效配置进入运行时
                if (!string.IsNullOrEmpty(item.Name))
                    list.Add(item);
            }
            return list;
        }

        private void EnsureTemplate()
        {
            if (File.Exists(_filePath))
                return;

            // 创建可直接使用的最小模板：
            // - 使能行预置开关值
            // - 三个给定行预置名称，地址留给用户填写
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("JogConfig");
            worksheet.Cell(1, 1).Value = "名称";
            worksheet.Cell(1, 2).Value = "地址id";
            worksheet.Cell(1, 3).Value = "数据类型";
            worksheet.Cell(1, 4).Value = "开状态";
            worksheet.Cell(1, 5).Value = "关状态";

            worksheet.Cell(2, 1).Value = "使能";
            worksheet.Cell(2, 3).Value = "1";
            worksheet.Cell(2, 4).Value = "0";

            worksheet.Cell(3, 1).Value = "电流给定";
            worksheet.Cell(4, 1).Value = "速度给定";
            worksheet.Cell(5, 1).Value = "位置给定";
            workbook.SaveAs(_filePath);
        }
    }
}
