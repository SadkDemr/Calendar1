using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Calendar1.Models
{
    public class EmployeeAttendance
    {
        public string EmployeeName { get; set; }
        public List<DateTime> AttendanceDates { get; set; } // Bu, çalışanın her bir katılım tarihini saklar.
        public DateTime StartDate { get; set; } // Bu, çalışanın işe başladığı tarihi saklar.
        public DateTime EndDate { get; set; } // Bu, çalışanın işi bittiği tarihi saklar.

        public string EmployeeTeam { get; set; }


        //public List<string> AvailableEmployeeNames { get; } = new List<string> { "Sadık", "Engin", "Emre", "Mücahit", "Gizem", "Furkan", "Zelal", "Birkan" };
        public List<string> AvailableEmployeeTeam { get; } = new List<string> { "Djital Uygulama", "Bankacılık" };
        public List<int> AvailableDeskNumbers { get; } = Enumerable.Range(1, 12).ToList();

        public void CalculateTeamAttendance(int year, int month, int day, List<EmployeeAttendance> allEmployeeAttendance)
        {
            var teams = AvailableEmployeeTeam.Distinct().ToList();

            foreach (var team in teams)
            {
                var teamAttendance = allEmployeeAttendance
                    .Where(e => e.EmployeeTeam == team)
                    .SelectMany(e => e.AttendanceDates)
                    .Count(d => d.Year == year && d.Month == month && d.Day == day);

                Console.WriteLine($"Takım: {team}, Tarih: {year}-{month}-{day}, Katılım: {teamAttendance}");
            }
        }
    }

}