using Calendar1.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Calendar1.Controllers
{
    public class ManagerController : Controller
    {
        private ExcelService _excelService;

        public ManagerController()
        {                                
            _excelService = new ExcelService("C:\\Users\\Sadik Demir\\Desktop\\calendar.xlsx");
        }
        // GET: Manager
        [HttpGet]
        public ActionResult ManagerCalendar()
        {
            // Ayların listesini oluşturun
            var months = CultureInfo.CurrentCulture.DateTimeFormat.MonthNames
                .Where(m => !string.IsNullOrEmpty(m)) // Boş olmayan ay isimlerini alır
                .Select((name, index) => new { Name = name, Value = index + 1 }) // Ay ismi ve sayısal değeri ile
                .ToList();

            ViewBag.Months = new SelectList(months, "Value", "Name");

            var managersFromExcel = ExcelService.GetEmployeeDetailsFromExcelManager("C:\\Users\\Sadik Demir\\Desktop\\calendar.xlsx"); // Doğru dosya yolu
            var model = new EmployeeDesk
            {
                Managers = managersFromExcel
            };
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ManagerCalendar(EmployeeDeskEventViewModel eventViewModel)
        {
            if (ModelState.IsValid)
            {
                // eventViewModel verilerini Excel'e kaydedin
                _excelService.AddDeskRecordManager(eventViewModel);

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
                bool success = _excelService.UpdateEventInExcelManage(updatedEvent);

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
            var events = _excelService.GetEventsFromExcelManager();

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
            bool isSuccess = _excelService.DeleteEventInExcelManager(eventViewModel);

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
                _excelService.GenerateSchedule(2023,selectedMonth);

                // Başarılı olduğunu belirten mesajı saklayın
                TempData["Success"] = "Ofis yerleştirme işlemi başarıyla tamamlandı.";
            }
            catch (Exception ex)
            {
                // Hata oluşursa kullanıcıya bir mesaj göster
                TempData["Error"] = "Bir hata oluştu: " + ex.Message;
            }

            // Kullanıcıyı AddDeskRecord sayfasına yönlendir
            return RedirectToAction("ManagerCalendar");
        }

        public ActionResult Delete(int selectedMonth)
        {
            Session["SelectedMonth"] = selectedMonth;
            _excelService.DeleteMonthManager(selectedMonth);

            return RedirectToAction("ManagerCalendar");
        }


    }
}