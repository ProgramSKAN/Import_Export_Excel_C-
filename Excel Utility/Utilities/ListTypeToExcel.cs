using Excel_Utility.Utilities.CustomStyleDeleteColumn;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace Excel_Utility.Utilities
{
    public static class ListTypeToExcel
    {
        public static byte[] Export<T>(IList<T> dataToExport,string workSheetName)
        {
            byte[] result = null;
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            using(ExcelPackage package=new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(workSheetName);
                int startIndex = 1;

                worksheet.Cells["A" + startIndex].LoadFromCollection(Collection:dataToExport,PrintHeaders:true,TableStyle:TableStyles.Light1);
                result = package.GetAsByteArray();
            }

            return result;
        }

        public static byte[] ExportWithCustomStyle<T>(IList<T> dataToExport, string workSheetName, List<KeyValuePair<string, string>> columnsToExport)
        {
            byte[] result = null;
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(workSheetName);
                int startIndex = 1;

                worksheet.Cells["A" + startIndex].LoadFromCollection(Collection: dataToExport, PrintHeaders: true);

                Custom.SetHeaderStyle(properties, worksheet, startIndex);
                Custom.SetColumnStyleAndHeader(columnsToExport, properties, worksheet, startIndex);
                Custom.SetCellStyle(dataToExport, properties, worksheet, startIndex);
                Custom.DeleteColumn(columnsToExport, properties, worksheet, startIndex);
                result = package.GetAsByteArray();
            }

            return result;
        }

    }
}
