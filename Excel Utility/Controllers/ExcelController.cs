using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Threading.Tasks;
using Excel_Utility.Models;
using Excel_Utility.Utilities;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace Excel_Utility.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {

        [HttpGet("scratch")]
        public object GetManualGeneratedExcel()
        {
            return File((byte[])ExcelUtilityFromScratch.Export(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Employee.xlsx");
        }


        [HttpGet("wrap1")]
        public object GetExcel()
        {
            List<Employee> employees= new List<Employee>();
            foreach(var i in Enumerable.Range(0, 50).ToList())
            {
                employees.Add(new Employee
                { Id=i,
                   Date=DateTime.Now,
                   Name="name"+i,
                   IsJoined=true,
                   DeleteIt="del"+i
                 });
            }

            dynamic employeesObj = new List<dynamic>();
            dynamic row = new ExpandoObject();

            foreach (var i in Enumerable.Range(0, 50).ToList())
            {
                row.Id = i;
                row.Date = DateTime.Now;
                row.Name = "name" + i;
                row.IsJoined = true;
                row.DeleteIt = "del" + i;
                employeesObj.Add(row);
            }

            List<KeyValuePair<string, string>>  columnsToExport = ColumnToExportEmployees();

            //byte[] result = ExcelUtilityWrapper.Export(countries, "CountriesList");
            //byte[] result = ListTypeToExcel.ExportWithCustomStyle(employees, "EmployeeList",columnsToExport);
            byte[] result = DyanmicListToExcel.ExportDynamicListObjectArray(employeesObj, "EmployeeList", columnsToExport);

            return File((byte[])result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Countries.xlsx");
        }



        private List<KeyValuePair<string, string>> ColumnToExportEmployees()
        {
            List<KeyValuePair<string, string>> columnsToExport = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("Id", "u_Id"),
                new KeyValuePair<string, string>("Date", "u_Date"),
                new KeyValuePair<string, string>("Name", "u_Name"),
                new KeyValuePair<string, string>("IsJoined","u_IsJoined"),
                //delete "DeleteIt" column
           };
            return columnsToExport;
        }

    }
}
