using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel_Utility.Models
{
    public class Employee
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public bool IsJoined { get; set; }
        public string DeleteIt { get; set; }
    }
}
