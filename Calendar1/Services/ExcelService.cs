using Calendar1.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

public class ExcelService
{
    private readonly string _excelPath;

    private const int MaxPeoplePerDay = 9;
    private const int RequiredDaysPerMonth = 7;
    private List<string> weekdays = new List<string> { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday" };


    //public IEnumerable<object> AvailableEmployeeTeam { get; private set; }
    public List<string> AvailableEmployeeTeam { get; } = new List<string> { "Djital Uygulama", "Bankacılık" };

    public ExcelService(string excelPath)
    {
        _excelPath = excelPath;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // EPPlus 5.0 ve sonrası için lisans bilgisi
    }

    public List<TeamAttendance> GetTeamAttendanceForDayMonthAndYear(int day, int month, int year)
    {
        List<TeamAttendance> teamAttendanceRecords = new List<TeamAttendance>();
        List<string> teamNames = new List<string>(); // Takım isimlerini saklamak için bir liste ekleyin

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                DateTime startDate;
                DateTime endDate;

                if (DateTime.TryParse(worksheet.Cells[row, 2].Text, out startDate)
                    && DateTime.TryParse(worksheet.Cells[row, 3].Text, out endDate)
                    && startDate.Day <= day && endDate.Day >= day
                    && startDate.Month == month && endDate.Month == month
                    && startDate.Year == year && endDate.Year == year)
                {
                    string teamName = worksheet.Cells[row, 4].Text;

                    if (!teamNames.Contains(teamName))
                    {
                        teamNames.Add(teamName);
                    }

                    TeamAttendance teamAttendance = teamAttendanceRecords.FirstOrDefault(t => t.TeamName == teamName);
                    if (teamAttendance == null)
                    {
                        teamAttendance = new TeamAttendance
                        {
                            TeamName = teamName,
                            Date = new DateTime(year, month, day),
                            EmployeeCount = 1
                        };
                        teamAttendanceRecords.Add(teamAttendance);
                    }
                    else
                    {
                        teamAttendance.EmployeeCount++;
                    }
                }
            }
        }

        // Takım isimlerini içeren listeye bakarak eksik takımları ekleyin
        foreach (var teamName in teamNames)
        {
            if (!teamAttendanceRecords.Any(t => t.TeamName == teamName))
            {
                teamAttendanceRecords.Add(new TeamAttendance
                {
                    TeamName = teamName,
                    Date = new DateTime(year, month, day),
                    EmployeeCount = 0
                });
            }
        }

        return teamAttendanceRecords;
    }

    public Dictionary<string, EmployeeAttendance> GetEmployeeAttendanceForMonthAndYear(int? month, int year)
    {
        Dictionary<string, EmployeeAttendance> attendanceRecords = new Dictionary<string, EmployeeAttendance>();

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                var employeeName = worksheet.Cells[row, 1].GetValue<string>();

                if (!attendanceRecords.ContainsKey(employeeName))
                {
                    attendanceRecords[employeeName] = new EmployeeAttendance
                    {
                        EmployeeName = employeeName,
                        AttendanceDates = new List<DateTime>(),
                        StartDate = DateTime.MaxValue,
                        EndDate = DateTime.MinValue
                    };
                }

                DateTime startDate, endDate;

                try
                {
                    startDate = worksheet.Cells[row, 2].GetValue<DateTime>();
                    endDate = worksheet.Cells[row, 3].GetValue<DateTime>();
                }
                catch
                {
                    // Tarih çözümlenemedi, bu satırı atlayın veya hata mesajını loglayın.
                    continue;
                }

                // Ay ve yıl eşleşiyorsa, katılım tarihlerini ekleyin.
                if ((month.HasValue && startDate.Month == month.Value && endDate.Month == month.Value && startDate.Year == year && endDate.Year == year)
                    || (!month.HasValue && startDate.Year == year && endDate.Year == year))
                {
                    if (startDate < attendanceRecords[employeeName].StartDate)
                    {
                        attendanceRecords[employeeName].StartDate = startDate;
                    }
                    if (endDate > attendanceRecords[employeeName].EndDate)
                    {
                        attendanceRecords[employeeName].EndDate = endDate;
                    }
                    for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
                    {
                        attendanceRecords[employeeName].AttendanceDates.Add(date);
                    }
                }
            }
        }

        return attendanceRecords;
    }









    public void AddDeskRecord(EmployeeDeskEventViewModel eventViewModel)
    {
        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];

            // Sonraki boş satırı bulma
            int row = 1;
            while (worksheet.Cells[row, 1].Text != string.Empty)
            {
                row++;
            }

            worksheet.Cells[row, 1].Value = eventViewModel.title; // Düzeltildi: EmployeeName -> title
                                                                  // Diğer ilgili verileri ekleyin, örneğin ekip adı ve tarih

            // Tarihi doğru biçimde ayarla:
            worksheet.Cells[row, 4].Value = eventViewModel.start.ToOADate(); // Düzeltildi: Date -> start
            worksheet.Cells[row, 4].Style.Numberformat.Format = "dd/MM/yyyy";

            package.Save();
        }
    }

    public List<EmployeeDeskEventViewModel> GetEventsFromExcel()
    {
        var events = new List<EmployeeDeskEventViewModel>();

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int row = 2; // İlk satır başlık satırı olduğundan, verilerin 2. satırdan başladığını varsayarak başlıyoruz.

            while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Text))
            {
                if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text) || string.IsNullOrWhiteSpace(worksheet.Cells[row, 3].Text))
                {
                    row++;
                    continue; // Eğer tarih hücreleri boşsa, bu satırı atla ve bir sonraki satıra geç.
                }

                try
                {
                    var eventViewModel = new EmployeeDeskEventViewModel
                    {
                        title = worksheet.Cells[row, 1].Text,
                        start = DateTime.ParseExact(worksheet.Cells[row, 2].Text, "dd.MM.yyyy", CultureInfo.InvariantCulture),
                        end = DateTime.ParseExact(worksheet.Cells[row, 3].Text, "dd.MM.yyyy", CultureInfo.InvariantCulture),
                        team = worksheet.Cells[row, 4].Text,
                    };

                    events.Add(eventViewModel);
                }
                catch (ArgumentException)
                {
                    Console.WriteLine($"Satır: {row}, Başlangıç Değeri: {worksheet.Cells[row, 2].Text}, Bitiş Değeri: {worksheet.Cells[row, 3].Text}");
                    // Geçerli olmayan tarih değeri yakalandı. Bu satırı atla.
                }

                row++;
            }
        }

        return events;
    }


    public bool UpdateEventInExcel(EmployeeDeskEventViewModel updatedEvent)
    {
        bool updated = false;

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int row = 2; // İlk satır başlık satırı olduğundan, verilerin 2. satırdan başladığını varsayarak başlıyoruz.

            // Son satırı bulmak için döngü
            while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Text))
            {
                row++;
            }

            try
            {
                worksheet.Cells[row, 1].Value = updatedEvent.title;
                worksheet.Cells[row, 2].Value = updatedEvent.start.ToString("dd.MM.yyyy");
                worksheet.Cells[row, 3].Value = updatedEvent.end.ToString("dd.MM.yyyy");
                worksheet.Cells[row, 4].Value = updatedEvent.team;  
                updated = true;
            }
            catch (ArgumentException)
            {
                Console.WriteLine($"Satır: {row}, Başlangıç Değeri: {worksheet.Cells[row, 2].Text}, Bitiş Değeri: {worksheet.Cells[row, 3].Text}");
                // Tarih değerleriyle ilgili bir sorun varsa, bu satırı atla.
                updated = false;
            }

            if (updated)
                package.Save();  // Değişiklikleri kaydet
        }

        return updated;
    }



    public bool DeleteEventInExcel(EmployeeDeskEventViewModel eventToDelete)
    {
        bool deleted = false;

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int row = 2;

            while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Text))
            {
                if (worksheet.Cells[row, 1].Text == eventToDelete.title &&
                    worksheet.Cells[row, 2].Text == eventToDelete.start.ToString("dd.MM.yyyy") &&
                    worksheet.Cells[row, 3].Text == eventToDelete.end.ToString("dd.MM.yyyy") &&
                    worksheet.Cells[row, 4].Text == eventToDelete.team) 
                {
                    worksheet.DeleteRow(row); // Satırı sil
                    deleted = true;
                    break;
                }
                row++;
            }

            if (deleted)
                package.Save(); // Değişiklikleri kaydet
        }

        return deleted;
    }



    public static List<Employee> GetEmployeeDetailsFromExcel(string path)
{
    List<Employee> employees = new List<Employee>();

    using (var pck = new OfficeOpenXml.ExcelPackage())
    {
        using (var stream = File.OpenRead(path))
        {
            pck.Load(stream);
        }

        var ws = pck.Workbook.Worksheets["Employees"];

        // Satırları kontrol et, 1. satır başlık olduğu için 2'den başla
        for (int row = 2; row <= ws.Dimension.End.Row; row++)
        {
            employees.Add(new Employee
            {
                Name = ws.Cells[row, 1].Text,
                Team = ws.Cells[row, 2].Text
            });
        }
    }
    return employees;
}

    public void AssignEmployeeDeskRandomly(int month)
    {
        // Load the Excel package
        var fileInfo = new FileInfo(_excelPath);
        using (var package = new ExcelPackage(fileInfo))
        {
            // Get the Employees sheet
            var employeeSheet = package.Workbook.Worksheets["Employees"];
            var scheduleSheet = package.Workbook.Worksheets["Sayfa1"] ?? package.Workbook.Worksheets.Add("Sayfa1");
            var employees = new List<EmployeeDeskEventViewModel>();
            int totalRows = employeeSheet.Dimension.End.Row;


            // ScheduleSheet preparation
            scheduleSheet.Cells[1, 1].Value = "Title";
            scheduleSheet.Cells[1, 2].Value = "Start";
            scheduleSheet.Cells[1, 3].Value = "End";
            scheduleSheet.Cells[1, 4].Value = "Team";


            // Delete existing records for the selected month
            for (int row = 2; row <= scheduleSheet.Dimension.End.Row; row++)
            {
                if (DateTime.TryParse(scheduleSheet.Cells[row, 2].Text, out DateTime startDate))
                {
                    if (startDate.Month == month && startDate.Year == DateTime.Now.Year)
                    {
                        scheduleSheet.DeleteRow(row);
                        row--; // Since we are deleting rows, reduce the row count
                    }
                }
            }

            int recordRow = scheduleSheet.Dimension.End.Row + 1;

            for (int row = 2; row <= totalRows; row++)
            {
                employees.Add(new EmployeeDeskEventViewModel
                {
                    title = employeeSheet.Cells[row, 1].Value.ToString(),
                    team = employeeSheet.Cells[row, 2].Value.ToString()
                });
            }


            // Randomly distribute employees over the month
            var daysInMonth = DateTime.DaysInMonth(DateTime.Now.Year, month);
            var random = new Random();

            // Her takım için günlük katılımı takip eden sözlük yapısı
            Dictionary<string, Dictionary<int, int>> teamDailyAttendance = new Dictionary<string, Dictionary<int, int>>();

            // Takımların listesini oluştur ve her gün için katılım sözlüğünü başlat
            foreach (var employee in employees)
            {
                if (!teamDailyAttendance.ContainsKey(employee.team))
                {
                    teamDailyAttendance[employee.team] = new Dictionary<int, int>();
                    for (int day = 1; day <= daysInMonth; day++)
                    {
                        var date = new DateTime(DateTime.Now.Year, month, day);
                        if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                        {
                            teamDailyAttendance[employee.team][day] = 0;
                        }
                    }
                }
            }

            // Her gün için çalışan sayısını takip eden bir sözlük yapısı oluştur.
            Dictionary<int, List<string>> dailyEmployeeAssignments = new Dictionary<int, List<string>>();

            // Aydaki tüm günler için sözlüğü başlat.
            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(DateTime.Now.Year, month, day);
                // Hafta sonları hariç her gün için boş liste ile başlar.
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    dailyEmployeeAssignments[day] = new List<string>();
                }
            }


            foreach (var employee in employees)
            {
                var daysAssigned = new HashSet<int>();
                int maxAttempts = 100;  // Her çalışan için maksimum deneme sayısı artırıldı.

                while (daysAssigned.Count < 7 && maxAttempts > 0)
                {
                    // Günlük atama sayısı 4'ün altında olan günleri önceliklendir
                    var prioritizedDays = dailyEmployeeAssignments
                        .Where(kv => kv.Value.Count < 4 && !daysAssigned.Contains(kv.Key))
                        .Select(kv => kv.Key)
                        .ToList();

                    int day;
                    if (prioritizedDays.Count > 0)
                    {
                        day = prioritizedDays[random.Next(prioritizedDays.Count)];
                    }
                    else
                    {
                        day = random.Next(1, daysInMonth + 1);
                    }

                    var date = new DateTime(DateTime.Now.Year, month, day);
                    if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday && !daysAssigned.Contains(day))
                    {
                        bool canAssign = true;

                        // Gün içinde en fazla 10 kişi olmasını kontrol et
                        if (dailyEmployeeAssignments[day].Count >= 8)
                        {
                            canAssign = false;
                        }

                        // Her takımdan en az 1 üyenin ofiste olmasını sağla
                        if (canAssign && teamDailyAttendance[employee.team][day] < 1)
                        {
                            canAssign = true;
                        }
                        else if (canAssign)
                        {
                            foreach (var team in teamDailyAttendance)
                            {
                                if (team.Key != employee.team && team.Value[day] < 1)
                                {
                                    canAssign = false;
                                    break;
                                }
                            }
                        }

                        if (canAssign)
                        {
                            dailyEmployeeAssignments[day].Add(employee.title);
                            daysAssigned.Add(day);
                            teamDailyAttendance[employee.team][day]++;
                        }
                    }

                    maxAttempts--;
                }

                // Eğer 7 gün doldurulamazsa, hata logu veya uygun bir işlem yapılabilir.
                if (daysAssigned.Count < 7)
                {
                    // Hata işleme veya loglama
                    Console.WriteLine($"Çalışan {employee.title} için yeterli gün sayısı atanamadı.");
                }
                else
                {
                    // Excel'e kaydetme işlemleri
                    foreach (var assignedDay in daysAssigned)
                    {
                        employee.start = new DateTime(DateTime.Now.Year, month, assignedDay);
                        employee.end = employee.start;

                        scheduleSheet.Cells[recordRow, 1].Value = employee.title;
                        scheduleSheet.Cells[recordRow, 2].Value = employee.start.ToString("dd.MM.yyyy");
                        scheduleSheet.Cells[recordRow, 3].Value = employee.end.ToString("dd.MM.yyyy");
                        scheduleSheet.Cells[recordRow, 4].Value = employee.team;
                        recordRow++;
                    }
                }
            }


            // Tüm atamalar yapıldıktan sonra her gün için en az 4 çalışanın olduğundan emin ol
            foreach (var day in Enumerable.Range(1, daysInMonth))
            {
                var date = new DateTime(DateTime.Now.Year, month, day);
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    while (dailyEmployeeAssignments[day].Count < 4)
                    {
                        var additionalEmployee = employees
                            .Where(e => !dailyEmployeeAssignments[day].Contains(e.title))
                            .OrderBy(_ => random.Next())
                            .FirstOrDefault();

                        if (additionalEmployee != null)
                        {
                            dailyEmployeeAssignments[day].Add(additionalEmployee.title);
                            teamDailyAttendance[additionalEmployee.team][day]++;
                        }
                        else
                        {
                            // Eğer ekstra çalışan yoksa, döngüden çık
                            break;
                        }
                    }
                }
            }



            // Save the Excel package
            package.Save();
        }
    }

    public void DeleteMonth(int selmonth)
    {
        // Load the Excel package
        var fileInfo = new FileInfo(_excelPath);
        using (var package = new ExcelPackage(fileInfo))
        {
            // Get the Employees sheet
            var employeeSheet = package.Workbook.Worksheets["Employees"];
            var scheduleSheet = package.Workbook.Worksheets["Sayfa1"] ?? package.Workbook.Worksheets.Add("Sayfa1");


            // ScheduleSheet preparation
            scheduleSheet.Cells[1, 1].Value = "Title";
            scheduleSheet.Cells[1, 2].Value = "Start";
            scheduleSheet.Cells[1, 3].Value = "End";
            scheduleSheet.Cells[1, 4].Value = "Team";


            // Delete existing records for the selected month
            for (int row = 2; row <= scheduleSheet.Dimension.End.Row; row++)
            {
                if (DateTime.TryParse(scheduleSheet.Cells[row, 2].Text, out DateTime startDate))
                {
                    if (startDate.Month == selmonth && startDate.Year == DateTime.Now.Year)
                    {
                        scheduleSheet.DeleteRow(row);
                        row--; // Since we are deleting rows, reduce the row count
                    }
                }
            }

            // Save changes to the Excel file
            package.Save();
        }
    }

    public void CreateMonthlyExcelFile(int selectedMonth, int year)
    {
        var fileInfo = new FileInfo(_excelPath);
        using (var package = new ExcelPackage(fileInfo))
        {
            // 'Sayfa1' sayfasını oku
            var sheet = package.Workbook.Worksheets["Sayfa1"];
            int totalRows = sheet.Dimension.End.Row;

            // Masaüstüne kaydetmek için dosya yolu oluştur
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = $"{selectedMonth}_{year}.xlsx";
            string fullPath = Path.Combine(desktopPath, fileName);

            // Yeni Excel dosyası için hazırlık
            //var newFile = new FileInfo($"{selectedMonth}_{year}.xlsx");
            using (var newPackage = new ExcelPackage(new FileInfo(fullPath)))
            {
                var newSheet = newPackage.Workbook.Worksheets.Add("Events");

                // Sütun başlıklarını ve çalışanları toplama
                var employeesByDate = new Dictionary<DateTime, List<string>>();
                var teamsByEmployee = new Dictionary<string, string>();

                for (int row = 2; row <= totalRows; row++)
                {
                    string title = sheet.Cells[row, 1].Value.ToString();
                    DateTime start = DateTime.ParseExact(sheet.Cells[row, 2].Value.ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture);
                    string team = sheet.Cells[row, 4].Value.ToString();

                    // Ay ve yıl kontrolü
                    if (start.Month == selectedMonth && start.Year == year)
                    {
                        if (!employeesByDate.ContainsKey(start))
                            employeesByDate[start] = new List<string>();
                        employeesByDate[start].Add(title);

                        if (!teamsByEmployee.ContainsKey(title))
                            teamsByEmployee[title] = team;
                    }
                }

                // Sütun başlıklarını ve verileri yazma
                int columnIndex = 1;
                foreach (var date in employeesByDate.Keys.OrderBy(d => d))
                {
                    newSheet.Cells[1, columnIndex].Value = date.ToString("dd MMMM yyyy dddd");

                    int rowIndex = 2;
                    foreach (var employee in employeesByDate[date])
                    {
                        newSheet.Cells[rowIndex, columnIndex].Value = employee;

                        var cell = newSheet.Cells[rowIndex, columnIndex];
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(GetTeamColor(teamsByEmployee[employee]));

                        rowIndex++;
                    }
                    columnIndex++;
                }

                newSheet.Cells.AutoFitColumns();
                newPackage.Save();
            }
        }
    }

    private System.Drawing.Color GetTeamColor(string team)
    {
        switch (team)
        {
            case "Bankacılık":
                return System.Drawing.Color.SkyBlue;
            case "Dijital Uygulama":
                return System.Drawing.Color.PaleVioletRed;
            default:
                return System.Drawing.Color.Black; // Varsayılan renk
        }
    }

    //Yonetciler Icin Olan Kısım
    public void AddDeskRecordManager(EmployeeDeskEventViewModel eventViewModel)
    {
        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets["Manager"];

            // Sonraki boş satırı bulma
            int row = 1;
            while (worksheet.Cells[row, 1].Text != string.Empty)
            {
                row++;
            }

            worksheet.Cells[row, 1].Value = eventViewModel.title; // Düzeltildi: EmployeeName -> title
                                                                  // Diğer ilgili verileri ekleyin, örneğin ekip adı ve tarih

            // Tarihi doğru biçimde ayarla:
            worksheet.Cells[row, 4].Value = eventViewModel.start.ToOADate(); // Düzeltildi: Date -> start
            worksheet.Cells[row, 4].Style.Numberformat.Format = "dd/MM/yyyy";

            package.Save();
        }
    }
    public List<EmployeeDeskEventViewModel> GetEventsFromExcelManager()
    {
        var events = new List<EmployeeDeskEventViewModel>();

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets["Manager"];
            int row = 2; // İlk satır başlık satırı olduğundan, verilerin 2. satırdan başladığını varsayarak başlıyoruz.

            while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Text))
            {
                if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text) || string.IsNullOrWhiteSpace(worksheet.Cells[row, 3].Text))
                {
                    row++;
                    continue; // Eğer tarih hücreleri boşsa, bu satırı atla ve bir sonraki satıra geç.
                }

                try
                {
                    var eventViewModel = new EmployeeDeskEventViewModel
                    {
                        title = worksheet.Cells[row, 1].Text,
                        start = DateTime.ParseExact(worksheet.Cells[row, 2].Text, "dd.MM.yyyy", CultureInfo.InvariantCulture),
                        end = DateTime.ParseExact(worksheet.Cells[row, 3].Text, "dd.MM.yyyy", CultureInfo.InvariantCulture),
                    };

                    events.Add(eventViewModel);
                }
                catch (ArgumentException)
                {
                    Console.WriteLine($"Satır: {row}, Başlangıç Değeri: {worksheet.Cells[row, 2].Text}, Bitiş Değeri: {worksheet.Cells[row, 3].Text}");
                    // Geçerli olmayan tarih değeri yakalandı. Bu satırı atla.
                }

                row++;
            }
        }

        return events;
    }

    public bool UpdateEventInExcelManage(EmployeeDeskEventViewModel updatedEvent)
    {
        bool updated = false;

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets["Manager"];
            int row = 2; // İlk satır başlık satırı olduğundan, verilerin 2. satırdan başladığını varsayarak başlıyoruz.

            // Son satırı bulmak için döngü
            while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Text))
            {
                row++;
            }

            try
            {
                worksheet.Cells[row, 1].Value = updatedEvent.title;
                worksheet.Cells[row, 2].Value = updatedEvent.start.ToString("dd.MM.yyyy");
                worksheet.Cells[row, 3].Value = updatedEvent.end.ToString("dd.MM.yyyy");
                updated = true;
            }
            catch (ArgumentException)
            {
                Console.WriteLine($"Satır: {row}, Başlangıç Değeri: {worksheet.Cells[row, 2].Text}, Bitiş Değeri: {worksheet.Cells[row, 3].Text}");
                // Tarih değerleriyle ilgili bir sorun varsa, bu satırı atla.
                updated = false;
            }

            if (updated)
                package.Save();  // Değişiklikleri kaydet
        }

        return updated;
    }

    public bool DeleteEventInExcelManager(EmployeeDeskEventViewModel eventToDelete)
    {
        bool deleted = false;

        using (var package = new ExcelPackage(new FileInfo(_excelPath)))
        {
            var worksheet = package.Workbook.Worksheets["Manager"];
            int row = 2;

            while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Text))
            {
                if (worksheet.Cells[row, 1].Text == eventToDelete.title &&
                    worksheet.Cells[row, 2].Text == eventToDelete.start.ToString("dd.MM.yyyy") &&
                    worksheet.Cells[row, 3].Text == eventToDelete.end.ToString("dd.MM.yyyy"))
                {
                    worksheet.DeleteRow(row); // Satırı sil
                    deleted = true;
                    break;
                }
                row++;
            }

            if (deleted)
                package.Save(); // Değişiklikleri kaydet
        }

        return deleted;
    }

    public static List<Manager> GetEmployeeDetailsFromExcelManager(string path)
    {
        List<Manager> managers = new List<Manager>();

        using (var pck = new OfficeOpenXml.ExcelPackage())
        {
            using (var stream = File.OpenRead(path))
            {
                pck.Load(stream);
            }

            var ws = pck.Workbook.Worksheets["ManagerName"];

            // Satırları kontrol et, 1. satır başlık olduğu için 2'den başla
            for (int row = 2; row <= ws.Dimension.End.Row; row++)
            {
                managers.Add(new Manager
                {
                    Name = ws.Cells[row, 1].Text
                });
            }
        }
        return managers;
    }

    public void GenerateSchedule(int year, int month)
    {
        FileInfo fileInfo = new FileInfo(_excelPath);
        using (var package = new ExcelPackage(fileInfo))
        {
            var namesSheet = package.Workbook.Worksheets["ManagerName"] ?? package.Workbook.Worksheets.Add("ManagerName");
            var scheduleSheet = package.Workbook.Worksheets["Manager"] ?? package.Workbook.Worksheets.Add("Manager");

            List<string> names = namesSheet.Cells["A2:A" + namesSheet.Dimension.End.Row]
                                .Select(c => c.Value.ToString())
                                .Distinct()
                                .ToList();

            var weekdays = GetAllWeekdaysOfMonth(year, month);
            Dictionary<DateTime, List<string>> schedule = new Dictionary<DateTime, List<string>>();


            // Delete existing records for the selected month
            for (int rows = 2; rows <= scheduleSheet.Dimension.End.Row; rows++)
            {
                if (DateTime.TryParse(scheduleSheet.Cells[rows, 2].Text, out DateTime startDate))
                {
                    if (startDate.Month == month && startDate.Year == DateTime.Now.Year)
                    {
                        scheduleSheet.DeleteRow(rows);
                        rows--; // Since we are deleting rows, reduce the row count
                    }
                }
            }

            // Her pazartesi için tüm isimlerin eklenmesi
            foreach (var monday in weekdays.Where(d => d.DayOfWeek == DayOfWeek.Monday))
            {
                schedule.Add(monday, new List<string>(names));
            }

            // Ay içindeki diğer günler için isimlerin ataması
            // Random gün seçimi ve isim atamaları
            Random rnd = new Random();
            var remainingDays = weekdays.Except(schedule.Keys).OrderBy(x => rnd.Next()).ToList();

            // Her ismin ayda 8 gün atandığından emin ol
            foreach (var name in names)
            {
                var assignedDays = schedule.Where(kvp => kvp.Value.Contains(name)).Select(kvp => kvp.Key).ToList();

                while (assignedDays.Count < 8)
                {
                    var day = remainingDays.FirstOrDefault(d => !schedule.ContainsKey(d) || !schedule[d].Contains(name));
                    if (day == default(DateTime)) break; // Eğer uygun gün kalmamışsa döngüden çık

                    if (schedule.ContainsKey(day))
                    {
                        schedule[day].Add(name);
                    }
                    else
                    {
                        schedule.Add(day, new List<string> { name });
                    }
                    assignedDays.Add(day);
                    remainingDays.Remove(day);
                }
            }

            // Schedule doldurma
            int row = 2; // 1. satır başlık satırı olduğu için 2'den başlatılıyor
            foreach (var entry in schedule.OrderBy(e => e.Key))
            {
                foreach (var name in entry.Value.Distinct())
                {
                    scheduleSheet.Cells[row, 1].Value = name;
                    scheduleSheet.Cells[row, 2].Value = entry.Key.ToString("dd.MM.yyyy");
                    scheduleSheet.Cells[row, 3].Value = entry.Key.ToString("dd.MM.yyyy");
                    row++;
                }
            }

            package.Save(); // Değişiklikleri kaydet
        }
    }

    private List<DateTime> GetAllWeekdaysOfMonth(int year, int month)
    {
        var dates = new List<DateTime>();
        DateTime startDate = new DateTime(year, month, 1);
        DateTime endDate = startDate.AddMonths(1).AddDays(-1);

        for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
        {
            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
            {
                dates.Add(date);
            }
        }

        return dates;
    }

    public void DeleteMonthManager(int selmonthManager)
    {
        // Load the Excel package
        var fileInfo = new FileInfo(_excelPath);
        using (var package = new ExcelPackage(fileInfo))
        {
            // Get the Employees sheet
            var employeeSheet = package.Workbook.Worksheets["ManagerName"];
            var scheduleSheet = package.Workbook.Worksheets["Manager"] ?? package.Workbook.Worksheets.Add("Manager");


            // ScheduleSheet preparation
            scheduleSheet.Cells[1, 1].Value = "Title";
            scheduleSheet.Cells[1, 2].Value = "Start";
            scheduleSheet.Cells[1, 3].Value = "End";


            // Delete existing records for the selected month
            for (int row = 2; row <= scheduleSheet.Dimension.End.Row; row++)
            {
                if (DateTime.TryParse(scheduleSheet.Cells[row, 2].Text, out DateTime startDate))
                {
                    if (startDate.Month == selmonthManager && startDate.Year == DateTime.Now.Year)
                    {
                        scheduleSheet.DeleteRow(row);
                        row--; // Since we are deleting rows, reduce the row count
                    }
                }
            }

            // Save changes to the Excel file
            package.Save();
        }
    }

    public void CreateMonthlyExcelFileManager(int selectedMonth, int year)
    {
        var fileInfo = new FileInfo(_excelPath);
        using (var package = new ExcelPackage(fileInfo))
        {
            // 'Manager' sayfasını oku
            var sheet = package.Workbook.Worksheets["Manager"];
            int totalRows = sheet.Dimension.End.Row;

            // Masaüstüne kaydetmek için dosya yolu oluştur
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = $"{selectedMonth}_{year}_Yöneticiler.xlsx";
            string fullPath = Path.Combine(desktopPath, fileName);

            using (var newPackage = new ExcelPackage(new FileInfo(fullPath)))
            {
                var newSheet = newPackage.Workbook.Worksheets.Add("Events");

                // Sütun başlıklarını ve çalışanları toplama
                var employeesByDate = new Dictionary<DateTime, List<string>>();

                for (int row = 2; row <= totalRows; row++)
                {
                    string title = sheet.Cells[row, 1].Value.ToString();
                    DateTime start = DateTime.ParseExact(sheet.Cells[row, 2].Value.ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture);

                    // Ay ve yıl kontrolü
                    if (start.Month == selectedMonth && start.Year == year)
                    {
                        if (!employeesByDate.ContainsKey(start))
                            employeesByDate[start] = new List<string>();
                        employeesByDate[start].Add(title);
                    }
                }

                // Sütun başlıklarını ve verileri yazma
                int columnIndex = 1;
                foreach (var date in employeesByDate.Keys.OrderBy(d => d))
                {
                    newSheet.Cells[1, columnIndex].Value = date.ToString("dd MMMM yyyy dddd");

                    int rowIndex = 2;
                    foreach (var employee in employeesByDate[date])
                    {
                        newSheet.Cells[rowIndex, columnIndex].Value = employee;
                        rowIndex++;
                    }
                    columnIndex++;
                }

                newSheet.Cells.AutoFitColumns();
                newPackage.Save();
            }
        }
    }

}