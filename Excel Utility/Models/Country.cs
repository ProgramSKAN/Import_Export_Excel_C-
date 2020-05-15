using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel_Utility.Models
{
    public class Country
    {
        public int CountryId { get; set; }
        public bool IsDeleted { get; set; }
        public string CountryAbbreviation { get; set; }
        public string CountryName { get; set; }
        public string CountryCallingCode { get; set; }
    }
}
