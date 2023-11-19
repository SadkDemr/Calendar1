using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Calendar1.Models
{
    public class EmployeeDeskEventViewModel
    {
        public string title { get; set; }
        public DateTime start { get; set; }
        public DateTime end { get; set; }
        public string team { get; set; }
        public HashSet<int> AssignedDays { get; set; } = new HashSet<int>(); // Atanmış günleri tutacak
    }
}
