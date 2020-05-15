using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;

namespace Excel_Utility.Utilities
{
    public static class ExcelUtilityFromScratch
    {
        public static byte[] Export()
        {
            byte[] result = null;
            var Employees = new[]
            {
                new{ Id=101,Name="Name1"},
                new{ Id=102,Name="Name2"},
                new{ Id=103,Name="Name3"},
                new{ Id=104,Name="Name4"},
                new{ Id=105,Name="Name5"},
                new{ Id=106,Name="Name6"},
            };
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("sheetName1");
            workSheet.TabColor = System.Drawing.Color.BlueViolet;

            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;
            workSheet.Row(1).Style.Font.Color.SetColor(Color.Brown);
            //workSheet.Row(1).Style.Fill.PatternType = ExcelFillStyle.LightGray;
            //workSheet.Row(1).Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
            //workSheet.Row(1).Style.Border.BorderAround(ExcelBorderStyle.Thick);

            workSheet.Cells[1, 1].Value = "S.No";
            workSheet.Cells[1, 2].Value = "Id";
            workSheet.Cells[1, 3].Value = "Name";

            int recordIndex = 2;
            foreach(var emp in Employees)
            {
                workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                workSheet.Cells[recordIndex, 2].Value = emp.Id;
                workSheet.Cells[recordIndex, 3].Value = emp.Name;
                //workSheet.Row(recordIndex).Style.Border.BorderAround(ExcelBorderStyle.Thin);
                recordIndex++;
            }

            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();

            workSheet.Cells[1, 1, 1, 5].Style.Fill.PatternType = ExcelFillStyle.LightGray;
            workSheet.Cells[1,1,1,5].Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
            using (ExcelRange r = workSheet.Cells[1,1,5,5])
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

            result = excel.GetAsByteArray();
            excel.Dispose();
            return result;
        }
    }
}
