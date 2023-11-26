using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Mvc;
using Calendar1.Models;
using OfficeOpenXml;
using System.IO;
using Newtonsoft.Json;

namespace Calendar1.Controllers
{
    public class HomeController : Controller
    {
        private ExcelService _excelService;

        public HomeController()
        {
            string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(projectDirectory, "calendar.xlsx");
            _excelService = new ExcelService(filePath);
        }

        [HttpGet]
        public ActionResult Index(string tableType, int? day, int? month, int? year)
        {
            if (!day.HasValue) day = DateTime.Now.Day;
            if (!month.HasValue) month = DateTime.Now.Month;
            if (!year.HasValue) year = DateTime.Now.Year;

            List<EmployeeAttendance> employeeAttendanceRecords = new List<EmployeeAttendance>();
            List<TeamAttendance> teamAttendanceRecords = new List<TeamAttendance>();

            try
            {
                var employeeAttendanceData = _excelService.GetEmployeeAttendanceForMonthAndYear(month.Value, year.Value);

                if (employeeAttendanceData != null)
                {
                    employeeAttendanceRecords = employeeAttendanceData.Values.ToList();
                }
                else
                {
                    ViewBag.ErrorMessage = "The specified month and year do not have any attendance records for employees.";
                }

                ViewBag.Day = day.Value;
                ViewBag.Year = year.Value;
                ViewBag.Month = month.HasValue ? month.Value : (int?)null;

                // İkinci tabloyu güncelle
                teamAttendanceRecords = _excelService.GetTeamAttendanceForDayMonthAndYear(day.Value, month.Value, year.Value);

                foreach (var teamAttendance in teamAttendanceRecords)
                {
                    teamAttendance.CalculateTeamAttendance(year.Value, month.Value, day.Value, employeeAttendanceRecords);
                }

                var model = new Tuple<List<EmployeeAttendance>, List<TeamAttendance>>(employeeAttendanceRecords, teamAttendanceRecords);

                return View(model);
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = "An error occurred while fetching attendance records: " + ex.Message;
            }

            return View();
        }

        [HttpGet]
        public ActionResult AddDeskRecord()
        {
            string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = Path.Combine(projectDirectory, "calendar.xlsx");
            // Ayların listesini oluşturun
            var months = CultureInfo.CurrentCulture.DateTimeFormat.MonthNames
                .Where(m => !string.IsNullOrEmpty(m)) // Boş olmayan ay isimlerini alır
                .Select((name, index) => new { Name = name, Value = index + 1 }) // Ay ismi ve sayısal değeri ile
                .ToList();

            ViewBag.Months = new SelectList(months, "Value", "Name");

            var employeesFromExcel = ExcelService.GetEmployeeDetailsFromExcel(filePath); // Doğru dosya yolu
            var model = new EmployeeDesk
            {
                Employees = employeesFromExcel
            };
            return View(model);
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult AddDeskRecord(EmployeeDeskEventViewModel eventViewModel)
        {
            if (ModelState.IsValid)
            {
                // eventViewModel verilerini Excel'e kaydedin
                _excelService.AddDeskRecord(eventViewModel);

                // Başarı durumunda JSON yanıtı döndürün
                return Json(new { success = true, message = "Etkinlik başarıyla eklendi." });
            }

            // Model geçerli değilse, hata mesajlarıyla birlikte JSON yanıtı döndürün
            return Json(new { success = false, message = "Etkinlik eklenirken bir hata oluştu." });
        }

        [HttpPost]
        public ActionResult UpdateEventInExcel(EmployeeDeskEventViewModel updatedEvent)
        {
            try
            {
                bool success = _excelService.UpdateEventInExcel(updatedEvent);

                return Json(new { success = success });
            }
            catch
            {
                return Json(new { success = false });
            }
        }


        [HttpGet]
        public ActionResult GetEvents()
        {
            var events = _excelService.GetEventsFromExcel();

            var jsonResult = Json(events, JsonRequestBehavior.AllowGet);
            jsonResult.MaxJsonLength = int.MaxValue; // Maksimum uzunluk sınırlamasını kaldırır

            // Tarih formatını ISO8601 uyumlu yap
            var jsonSettings = new JsonSerializerSettings
            {
                DateFormatHandling = DateFormatHandling.IsoDateFormat
            };

            return Content(JsonConvert.SerializeObject(events, jsonSettings), "application/json");
        }


        [HttpPost]
        public ActionResult DeleteEventInExcel(EmployeeDeskEventViewModel eventViewModel)
        {
            bool isSuccess = _excelService.DeleteEventInExcel(eventViewModel);

            if (isSuccess)
            {
                return Json(new { success = true, message = "Olay başarıyla silindi." });
            }
            else
            {
                return Json(new { success = false, message = "Olay silinirken bir hata oluştu." });
            }
        }


        [HttpPost]
        public ActionResult AssignDesks(int selectedMonth)
        {
            try
            {
                // Seçilen ayı session'da saklayın
                Session["SelectedMonth"] = selectedMonth;

                // Çalışanların ofis yerleşimini rastgele atayın
                _excelService.AssignEmployeeDeskRandomly(selectedMonth);

                // Başarılı olduğunu belirten mesajı saklayın
                TempData["Success"] = "Ofis yerleştirme işlemi başarıyla tamamlandı.";
            }
            catch (Exception ex)
            {
                // Hata oluşursa kullanıcıya bir mesaj göster
                TempData["Error"] = "Bir hata oluştu: " + ex.Message;
            }

            // Kullanıcıyı AddDeskRecord sayfasına yönlendir
            return RedirectToAction("AddDeskRecord");
        }

        public ActionResult Delete (int selectedMonth)
        {
            Session["SelectedMonth"] = selectedMonth;
            _excelService.DeleteMonth(selectedMonth);

            return RedirectToAction("AddDeskRecord");
        }



        [HttpPost]
        public ActionResult ExportEventsToExcel(int selectedMonth)
        {
            try
            {
             
                int year = DateTime.Now.Year;
                string monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(selectedMonth);

                // Excel dosyası oluşturma işlemi
                _excelService.CreateMonthlyExcelFile(selectedMonth, year);

                // Başarılı mesajı dön
                return RedirectToAction("AddDeskRecord");
            }
            catch (Exception ex)
            {
                // Hata durumunda hata mesajını dön
                return Json(new { success = false, message = "Hata: " + ex.Message });
            }
        }


    }

}
