using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Calendar1.Models
{
    public class TeamAttendance
    {
        public string TeamName { get; set; }
        public DateTime Date { get; set; }
        public int EmployeeCount { get; set; }


        // Her takım için ofiste olan kişi sayısını saklamak için bir liste ekleyin
        public List<EmployeeTeamAttendance> EmployeeTeamAttendances { get; set; } = new List<EmployeeTeamAttendance>();

        public void CalculateTeamAttendance(int year, int month, int day, List<EmployeeAttendance> allEmployeeAttendance)
        {
            var teamAttendance = allEmployeeAttendance
                .Where(e => e.EmployeeTeam == TeamName)
                .SelectMany(e => e.AttendanceDates)
                .Count(d => d.Year == year && d.Month == month && d.Day == day);

            // EmployeeCount'u sıfırlamadan önce güncel değerleri kullanarak ayarlayın
            EmployeeCount += teamAttendance;

            // Her takım için ofiste olan kişi sayısını hesaplayın ve saklayın
            EmployeeTeamAttendances.Clear();

            foreach (var employeeAttendance in allEmployeeAttendance)
            {
                if (employeeAttendance.EmployeeTeam == TeamName)
                {
                    int attendanceCount = employeeAttendance.AttendanceDates.Count(d => d.Year == year && d.Month == month && d.Day == day);
                    EmployeeTeamAttendance teamAttendanceInfo = new EmployeeTeamAttendance
                    {
                        EmployeeName = employeeAttendance.EmployeeName,
                        AttendanceCount = attendanceCount
                    };
                    EmployeeTeamAttendances.Add(teamAttendanceInfo);
                }
            }
        }

    }

    // Her takım için ofiste olan kişi sayısı ve çalışan isimlerini saklamak için bir sınıf ekleyin
    public class EmployeeTeamAttendance
    {
        public string EmployeeName { get; set; } // Çalışan Adı
        public int AttendanceCount { get; set; } // Ofiste Bulunma Sayısı
    }



}