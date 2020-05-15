using System;
using System.Collections.Generic;
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



    }
}
