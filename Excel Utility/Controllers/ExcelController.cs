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


        [HttpGet("wrap1")]
        public object GetExcel()
        {
            List<Country> countries= new List<Country>();
            foreach(var i in Enumerable.Range(0, 50).ToList())
            {
                countries.Add(new Country { 
                    CountryId = i, 
                    IsDeleted = false, 
                    CountryAbbreviation = "abb"+i, 
                    CountryName = "country"+i, 
                    CountryCallingCode = "10"+i });
            }
            

            return File((byte[])ExcelUtilityWrapper.Export(countries,"CountriesList"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Countries.xlsx");
        }

    }
}
