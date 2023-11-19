using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Calendar1.Models
{
    public class EmployeeDesk
    {
        public string EmployeeName { get; set; }
        public string EmployeeTeam{ get; set; }
        //public int DeskNumber { get; set; }
        public DateTime Date { get; set; }

        public List<string> AvailableEmployeeNames { get; set; }
        public List<string> AvailableEmployeeTeams { get; set; }

        public List<Employee> Employees { get; set; } = new List<Employee>();
        public List<Manager> Managers { get; set; } = new List<Manager>();


        // Eklemeler:
        //public List<string> AvailableEmployeeNames { get; } = new List<string> { "Sadık", "Engin", "Emre","Mücahit","Gizem","Furkan","Zelal","Birkan" }; // Buraya diğer isimleri ekleyebilirsiniz
        //public List<string> AvailableEmployeeTeam { get; } = new List<string> { "Djital Uygulama", "Bankacılık" };
        public List<int> AvailableDeskNumbers { get; } = Enumerable.Range(1, 12).ToList();
    }
}