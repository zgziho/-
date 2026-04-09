using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Word;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using WpfApp1.Models;
using WpfApp1.Service.Interfaces;

namespace WpfApp1.Service
{
    public  class ExcelService 
    {
            private readonly string _filePath;

            public ExcelService(string filePath = "Pars.xlsx")
            {
                _filePath = filePath;
            }


        // 另存为Excle
        public void SaveData(ObservableCollection<Pars> data, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Pars");

                worksheet.Cell(1, 1).Value = "序号";
                worksheet.Cell(1, 2).Value = "地址ID";
                worksheet.Cell(1, 3).Value = "数据名称";
                worksheet.Cell(1, 4).Value = "数据值";
                worksheet.Cell(1, 5).Value = "单位";
                worksheet.Cell(1, 6).Value = "数据类型";
                worksheet.Cell(1, 7).Value = "属性";
                worksheet.Cell(1, 8).Value = "数据范围";
                worksheet.Cell(1, 9).Value = "显示方式";
                worksheet.Cell(1, 10).Value = "数据系数";
                worksheet.Cell(1, 11).Value = "显示小数位";
                worksheet.Cell(1, 12).Value = "显示";
                worksheet.Cell(1, 13).Value = "背景颜色";
                worksheet.Cell(1, 14).Value = "数据说明";

                int row = 2;
                foreach (var p in data)
                {
                    worksheet.Cell(row, 1).Value = p.Id ?? "";
                    worksheet.Cell(row, 2).Value = p.ParsID ?? "";
                    worksheet.Cell(row, 3).Value = p.ParsNM ?? "";
                    worksheet.Cell(row, 4).Value = p.ParsVA ?? "";
                    worksheet.Cell(row, 5).Value = p.ParsDW ?? "";
                    worksheet.Cell(row, 6).Value = p.ParsLX ?? "";
                    worksheet.Cell(row, 7).Value = p.ParsSX ?? "";
                    worksheet.Cell(row, 8).Value = p.ParsFW ?? "";
                    worksheet.Cell(row, 9).Value = p.ParsXSFS ?? "";
                    worksheet.Cell(row, 10).Value = p.ParsXS ?? "";
                    worksheet.Cell(row, 11).Value = p.ParsXSW ?? "";
                    worksheet.Cell(row, 12).Value = p.ParsLB ?? "";
                    worksheet.Cell(row, 13).Value = p.ParsYS ?? "";
                    worksheet.Cell(row, 14).Value = p.ParsBZ ?? "";
                    row++;
                }
                workbook.SaveAs(filePath);
            }
        }

        // 从 Excel 加载数据
        public ObservableCollection<Pars> LoadData()
            {
                var list = new ObservableCollection<Pars>();
                if (!File.Exists(_filePath))
                    return list;
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    // 第一行是标题，数据从第二行开始
                    var rows = worksheet.RowsUsed().Skip(1);
                    foreach (var row in rows)
                    {
                        var pars = new Pars
                        {
                            Id = row.Cell(1).GetString(),
                            ParsID = row.Cell(2).GetString(),
                            ParsNM = row.Cell(3).GetString(),
                            ParsVA = row.Cell(4).GetString(),
                            ParsDW = row.Cell(5).GetString(),
                            ParsLX = row.Cell(6).GetString(),
                            ParsSX = row.Cell(7).GetString(),
                            ParsFW = row.Cell(8).GetString(),
                            ParsXSFS = row.Cell(9).GetString(),
                            ParsXS = row.Cell(10).GetString(),
                            ParsXSW = row.Cell(11).GetString(),
                            ParsLB = row.Cell(12).GetString(),
                            ParsYS = row.Cell(13).GetString(),
                            ParsBZ = row.Cell(14).GetString()
                        };
                        list.Add(pars);
                    }
                }
            }
            catch
            {
                MessageBox.Show("文件被打开");
            }
                return list;
            }

            // 保存数据到 Excel
        public void SaveData(ObservableCollection<Pars> pers)
            {
            try
            {

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Pars");
                    // 设置标题行（中文标题，方便查看）
                    worksheet.Cell(1, 1).Value = "序号";
                    worksheet.Cell(1, 2).Value = "地址ID";
                    worksheet.Cell(1, 3).Value = "数据名称";
                    worksheet.Cell(1, 4).Value = "数据值";
                    worksheet.Cell(1, 5).Value = "单位";
                    worksheet.Cell(1, 6).Value = "数据类型";
                    worksheet.Cell(1, 7).Value = "属性";
                    worksheet.Cell(1, 8).Value = "数据范围";
                    worksheet.Cell(1, 9).Value = "显示方式";
                    worksheet.Cell(1, 10).Value = "数据系数";
                    worksheet.Cell(1, 11).Value = "显示小数位";
                    worksheet.Cell(1, 12).Value = "显示";
                    worksheet.Cell(1, 13).Value = "背景颜色";
                    worksheet.Cell(1, 14).Value = "数据说明";

                    int row = 2;
                    foreach (var p in pers)
                    {
                        worksheet.Cell(row, 1).Value = p.Id ?? "";
                        worksheet.Cell(row, 2).Value = p.ParsID ?? "";
                        worksheet.Cell(row, 3).Value = p.ParsNM ?? "";
                        worksheet.Cell(row, 4).Value = "";
                        worksheet.Cell(row, 5).Value = p.ParsDW ?? "";
                        worksheet.Cell(row, 6).Value = p.ParsLX ?? "";
                        worksheet.Cell(row, 7).Value = p.ParsSX ?? "";
                        worksheet.Cell(row, 8).Value = p.ParsFW ?? "";
                        worksheet.Cell(row, 9).Value = p.ParsXSFS ?? "";
                        worksheet.Cell(row, 10).Value = p.ParsXS ?? "";
                        worksheet.Cell(row, 11).Value = p.ParsXSW ?? "";
                        worksheet.Cell(row, 12).Value = p.ParsLB ?? "";
                        worksheet.Cell(row, 13).Value = p.ParsYS ?? "";
                        worksheet.Cell(row, 14).Value = p.ParsBZ ?? "";
                        row++;
                    }
                    workbook.SaveAs(_filePath);
                }
            }
            catch
            {
              
            }
            }
        }
    }

