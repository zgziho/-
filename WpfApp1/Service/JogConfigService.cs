using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using WpfApp1.Models;

namespace WpfApp1.Service
{
    public class JogConfigService
    {
        private const string ResourceName = "WpfApp1.JogConfig.xlsx";

        public JogConfigService()
        {
        }

        public IReadOnlyList<JogConfigItem> Load()
        {
            var assembly = Assembly.GetExecutingAssembly();
            using var stream = assembly.GetManifestResourceStream(ResourceName) 
                ?? throw new FileNotFoundException("嵌入资源 JogConfig.xlsx 未找到", ResourceName);
            
            var list = new List<JogConfigItem>();
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RowsUsed().Skip(1);
            foreach (var row in rows)
            {
                var item = new JogConfigItem
                {
                    Name = row.Cell(1).GetString().Trim(),
                    AddressId = row.Cell(2).GetString().Trim(),
                    DataType = row.Cell(3).GetString().Trim(),
                    OnState = row.Cell(4).GetString().Trim(),
                    OffState = row.Cell(5).GetString().Trim()
                };
                if (!string.IsNullOrEmpty(item.Name))
                    list.Add(item);
            }
            return list;
        }
    }
}
