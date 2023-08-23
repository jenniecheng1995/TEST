using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;

namespace GISFCU.EIP4.HR.Helpers
{
    /// <summary>
    /// Epplus套件幫助類
    /// </summary>
    public class EpplusHelper
    {
        public static MemoryStream CreateExcel<T>(string sheetName, List<T> dataList, List<string> selector = null, List<EpplusFormat> Format = null) where T : class
        {
            PropertyInfo[] properties = null;
            MemoryStream stream = new MemoryStream();

            if (dataList.Count > 0)
            {
                Type type = dataList[0].GetType();

                properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);

                if (selector != null && selector.Count > 0)
                {
                    properties = properties.Where(x => selector.Contains(x.Name)).ToArray();
                }
                else
                {
                    properties = properties.ToArray();
                }

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    using (var range = worksheet.Cells[1, 1, 1, properties.Length])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(69, 146, 173));
                        range.Style.Font.Color.SetColor(Color.White);
                    }
                    int row = 1, col;
                    object objColValue;

                    for (int j = 0; j < properties.Length; j++)
                    {
                        row = 1;
                        col = j + 1;
                        var propertyName = properties[j].Name;
                        var displayNameAtt = properties[j].GetCustomAttribute(typeof(DisplayNameAttribute)) as DisplayNameAttribute;
                        worksheet.Cells[row, col].Value = displayNameAtt == null ? propertyName : displayNameAtt.DisplayName;
                    }

                    worksheet.View.FreezePanes(row + 1, 1);

                    for (int i = 0; i < dataList.Count; i++)
                    {
                        row = i + 2;
                        for (int j = 0; j < properties.Length; j++)
                        {
                            col = j + 1;
                            objColValue = properties[j].GetValue(dataList[i], null);
                            worksheet.Cells[row, col].Value = objColValue;
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns(10, 60);
                    worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    if (Format != null)
                    {
                        var startRow = 1;
                        var endRow = worksheet.Dimension.End.Row;

                        foreach (var item in Format)
                        {
                            var cells = worksheet.Cells[startRow, (int)item.StartCol, endRow, (int)item.EndCol];

                            foreach (var format in item.Format)
                            {
                                if (format.Type == "HorizontalAlignment")
                                    cells.Style.HorizontalAlignment = format.Setting;
                                if (format.Type == "Numberformat")
                                    cells.Style.Numberformat.Format = format.Setting;
                                if (format.Type == "WrapText")
                                    cells.Style.WrapText = true;
                            }
                        }
                    }

                    package.SaveAs(stream);
                    stream.Position = 0;
                }
            }
            else
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    worksheet.Cells[1, 1].Value = "無資料";

                    package.SaveAs(stream);
                    stream.Position = 0;
                }
            }
            return stream;
        }

        public static MemoryStream CreateSelExcel<T>(string sheetName, List<T> dataList, List<EpplusSelector> selector = null, List<EpplusFormat> Format = null) where T : class
        {
            PropertyInfo[] properties = null;

            MemoryStream stream = new MemoryStream();

            if (dataList.Count > 0)
            {
                Type type = dataList[0].GetType();

                properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);

                if (selector != null && selector.Count > 0)
                {
                    properties = properties.Where(x => selector.Select(y => y.Field).Contains(x.Name)).ToArray();
                }
                else
                {
                    properties = properties.ToArray();
                }

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    var selLength = selector.Count;
                    var propLength = properties.Length;

                    using (var range = worksheet.Cells[1, 1, 1, selLength])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(69, 146, 173));
                        range.Style.Font.Color.SetColor(Color.White);
                    }

                    int row = 1, col;

                    for (int j = 0; j < selLength; j++)
                    {
                        row = 1;
                        col = j + 1;
                        worksheet.Cells[row, col].Value = selector[j].Name;

                        if (selector[j].HeaderColor != null)
                            worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(selector[j].HeaderColor);
                    }

                    worksheet.View.FreezePanes(row + 1, 1);

                    for (int i = 0; i < dataList.Count; i++)
                    {
                        row = i + 2;

                        for (int j = 0; j < propLength; j++)
                        {
                            worksheet.Cells[row, j + 1].Value = properties[j].GetValue(dataList[i], null);
                        }
                        for (int j = propLength; j < selLength; j++)
                        {
                            var formula = selector[j].Formula;

                            if (!string.IsNullOrEmpty(selector[j].Formula))
                            {
                                var colChar = ExcelCellAddress.GetColumnLetter(j);
                                worksheet.Cells[row, j + 1].Formula = formula.Replace("{row}", row.ToString());
                            }
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns(10, 60);
                    worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    if (Format != null)
                    {
                        foreach (var item in Format)
                        {
                            var startRow = item.StartRow ?? 1;
                            var startCol = item.StartCol ?? 1;
                            var endRow = item.EndRow ?? worksheet.Dimension.End.Row;
                            var endCol = item.EndCol ?? worksheet.Dimension.End.Column;
                            var cells = worksheet.Cells[startRow, startCol, endRow, endCol];

                            foreach (var format in item.Format)
                            {
                                if (format.Type == "HorizontalAlignment")
                                    cells.Style.HorizontalAlignment = format.Setting;
                                if (format.Type == "Numberformat")
                                    cells.Style.Numberformat.Format = format.Setting;
                                if (format.Type == "WrapText")
                                    cells.Style.WrapText = true;
                            }
                        }
                    }

                    package.SaveAs(stream);
                    stream.Position = 0;
                }
            }
            else
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    worksheet.Cells[1, 1].Value = "無資料";

                    package.SaveAs(stream);
                    stream.Position = 0;
                }
            }
            return stream;
        }

        public static MemoryStream CreateExcelNoHeader<T>(string sheetName, List<T> dataList, List<string> selector = null) where T : class
        {
            PropertyInfo[] properties = null;
            MemoryStream stream = new MemoryStream();

            if (dataList.Count > 0)
            {
                Type type = dataList[0].GetType();

                properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);

                if (selector != null && selector.Count > 0)
                {
                    properties = properties.Where(x => selector.Contains(x.Name)).ToArray();
                }
                else
                {
                    properties = properties.ToArray();
                }

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    int row = 1, col;
                    object objColValue;
                    string colValue;

                    for (int i = 0; i < dataList.Count; i++)
                    {
                        row = i + 1;
                        for (int j = 0; j < properties.Length; j++)
                        {
                            col = j + 1;
                            objColValue = properties[j].GetValue(dataList[i], null);
                            colValue = objColValue == null ? "" : objColValue.ToString();
                            worksheet.Cells[row, col].Value = colValue;
                        }
                    }
                    package.SaveAs(stream);
                    stream.Position = 0;
                }
            }
            return stream;
        }
    }

    public class EpplusSelector
    {
        public string Field { get; set; }
        public dynamic Name { get; set; }
        public string Formula { get; set; }
        public dynamic HeaderColor { get; set; }
    }

    public class EpplusFormat
    {
        public int? StartCol { get; set; }
        public int? EndCol { get; set; }
        public int? StartRow { get; set; }
        public int? EndRow { get; set; }
        public List<Format> Format { get; set; }
    }

    public class Format
    {
        public string Type { get; set; }
        public dynamic Setting { get; set; }
    }
}