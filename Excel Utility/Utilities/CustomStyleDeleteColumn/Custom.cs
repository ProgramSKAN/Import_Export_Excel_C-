using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace Excel_Utility.Utilities.CustomStyleDeleteColumn
{
    public static class Custom
    {
        public static void SetHeaderStyle(PropertyDescriptorCollection properties, ExcelWorksheet worksheet,int startIndex)
        {
            using(ExcelRange r = worksheet.Cells[startIndex, 1, startIndex, properties.Count])
            {
                r.Style.Font.Color.SetColor(System.Drawing.Color.Brown);
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                r.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //for(int i = 1; i<= r.Columns; i++)
                //{
                //    worksheet.Column(i).AutoFit();
                //}
                
            }
        }

        //for setting column header names and column types (like datetime,...) 
        public static void SetColumnStyleAndHeader(List<KeyValuePair<string, string>> columnsToExport, PropertyDescriptorCollection properties, ExcelWorksheet workSheet, int startIndex)
        {
            for (int columnNumber = 1; columnNumber <= properties.Count; columnNumber++)
            {
                ExcelRange r = workSheet.Cells[startIndex, 1, startIndex, properties.Count];
                var current = columnsToExport.Find(p => p.Key.ToUpper() == r[1, columnNumber].Value.ToString().ToUpper());
                if (current.Value != null)
                {
                    r[1, columnNumber].Value = current.Value;
                }

                if (properties[columnNumber - 1].PropertyType == typeof(DateTime) || properties[columnNumber - 1].PropertyType == typeof(DateTime?))
                {
                    workSheet.Column(columnNumber).Style.Numberformat.Format = "dddd dd-mmm-yyyy h\"hours\"mm:ss AM/PM";
                }

                workSheet.Column(columnNumber).AutoFit();
            }
        }

        public static void SetCellStyle<T>(IList<T> dataToExport, PropertyDescriptorCollection properties, ExcelWorksheet workSheet, int startIndex)
        {
            using (ExcelRange r = workSheet.Cells[startIndex, 1, startIndex + dataToExport.Count, properties.Count])
            {
                r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                r.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                r.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            }
        }
        public static void DeleteColumn(List<KeyValuePair<string, string>> columnsToExport, PropertyDescriptorCollection properties, ExcelWorksheet workSheet, int startIndex)
        {
            List<int> columnsToDelete = new List<int>();
            for (int i = properties.Count - 1; i >= 0; i--)
            {
                ExcelRange r = workSheet.Cells[startIndex, 1, startIndex, properties.Count];
                if (r[1, i + 1].Value != null)
                {
                    var current = columnsToExport.Find(p => p.Value.ToUpper() == r[1, i + 1].Value.ToString().ToUpper());
                    if (current.Key == null)
                    {
                        columnsToDelete.Add(i + 1);
                    }
                }
            }

            foreach (int current in columnsToDelete)
            {
                workSheet.DeleteColumn(current);
            }
        }
    }
}
