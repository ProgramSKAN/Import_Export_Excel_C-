using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel_Utility.Utilities
{
    public static class DyanmicListToExcel
    {
        public static byte[] ExportDynamicListObjectArray<T>(IList<T> dataToExport, string workSheetName, List<KeyValuePair<string, string>> columnsToExport)
        {
            List<object[]> excelObject = new List<object[]>();

            var collection = dataToExport as IList<dynamic>;

            var dynamicObj = ((IDictionary<string, object>)collection.First()).ToDictionary(x => x.Key, x => x.Value);

            object[] headerArray = new object[1000];
            int count = 0;

            foreach (var item in columnsToExport)
            {
                headerArray[count++] = item.Value;
            }

            excelObject.Add(headerArray);

            foreach (var item in collection.ToList())
            {
                int incrementor = 0;
                object[] valueArray = new object[1000];
                var obj = (IDictionary<string, object>)item;

                foreach (var prop in dynamicObj)
                {
                    if (obj.ContainsKey(prop.Key))
                    {
                        if (obj[prop.Key] is DateTime  /*prop.Key == "Date"*/)
                        {
                            valueArray[incrementor++] = Convert.ToDateTime(obj[prop.Key]).ToString("dd-MMM-yyyy");
                        }
                        else
                        {
                            valueArray[incrementor++] = obj[prop.Key];
                        }
                    }
                }

                excelObject.Add(valueArray);
            }

            byte[] result = null;

            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(workSheetName);
                int startIndex = 1;

                workSheet.Cells["A" + startIndex].LoadFromArrays(excelObject);

                SetHeaderStyle(columnsToExport.Count, workSheet, startIndex);
                SetColumnStyleAndHeader(columnsToExport, columnsToExport.Count, workSheet, startIndex);
                SetCellStyle(dataToExport.Count, columnsToExport.Count, workSheet, startIndex);

                result = package.GetAsByteArray();
            }

            return result;
        }

        private static void SetHeaderStyle(int propertiesCount, ExcelWorksheet workSheet, int startIndex)
        {
            using (ExcelRange r = workSheet.Cells[startIndex, 1, startIndex, propertiesCount])
            {
                r.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                r.Style.Font.Bold = true;
                r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DimGray);
                r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                r.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }
        }
        private static void SetColumnStyleAndHeader(List<KeyValuePair<string, string>> columnsToExport, int propertiesCount, ExcelWorksheet workSheet, int startIndex)
        {
            for (int columnNumber = 1; columnNumber <= propertiesCount; columnNumber++)
            {
                ExcelRange r = workSheet.Cells[startIndex, 1, startIndex, propertiesCount];
                var current = columnsToExport.Find(p => p.Key.ToUpper() == r[1, columnNumber].Value.ToString().ToUpper());
                if (current.Value != null)
                {
                    r[1, columnNumber].Value = current.Value;
                }

                workSheet.Column(columnNumber).AutoFit();
            }

            workSheet.Column(1).Style.Numberformat.Format = "dd-mmm-yyyy";
        }
        private static void SetCellStyle(int dataToExportCount, int propertiesCount, ExcelWorksheet workSheet, int startIndex)
        {
            using (ExcelRange r = workSheet.Cells[startIndex, 1, startIndex + dataToExportCount, propertiesCount])
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
    }
}
