using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace Excel_Utility.Utilities
{
    public static class ExcelUtilityWrapper
    {
        public static byte[] Export<T>(IList<T> dataToExport,string workSheetName,List<KeyValuePair<string,string>> columnsToExport=null)
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
    }
}
