using IcsCalendar2Excel.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Ical.Net;
using System.Net.Http;
using System.IO;
using Microsoft.Extensions.Options;

namespace IcsCalendar2Excel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IOptions<CalendarSettings> _calendarSettings;

        public HomeController(ILogger<HomeController> logger, IOptions<CalendarSettings> calendarSettings)
        {
            _logger = logger;
            _calendarSettings = calendarSettings;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        public async Task<IActionResult> GenerateCalendar(CalendarInputModel input)
        {
            if (input.Urls == null || !input.Urls.Any())
            {
                ModelState.AddModelError("Urls", "At least one URL is required.");
                return View("Index", input);
            }

            var calendars = new List<Calendar>();
            using var httpClient = new HttpClient();

            foreach (var url in input.Urls)
            {
                var response = await httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsStringAsync();
                var calendar = Calendar.Load(content);
                calendars.Add(calendar);
            }

            var excelFile = GenerateExcelFile(input.Year, calendars);
            var output = new CalendarOutputModel
            {
                FileName = $"Calendar_{input.Year}.xlsx",
                FileContent = excelFile
            };

            return View("Download", output);
        }

        // 1. Try
        ////private byte[] GenerateExcelFile(int year, List<Calendar> calendars)
        ////{
        ////    using var workbook = new XLWorkbook();
        ////    var ws = workbook.AddWorksheet($"{year} Calendar");

        ////    // Add your logic to populate the worksheet with the calendar and ICS data.

        ////    using var ms = new MemoryStream();
        ////    workbook.SaveAs(ms);
        ////    return ms.ToArray();
        ////}

        //// 2. try
        ////private byte[] GenerateExcelFile(int year, List<Calendar> calendars)
        ////{
        ////    using var workbook = new XLWorkbook();
        ////    var ws = workbook.AddWorksheet($"{year} Calendar");

        ////    int startRow = 1;
        ////    int currentRow = startRow;
        ////    int currentColumn = 1;

        ////    for (int halfYear = 1; halfYear <= 2; halfYear++)
        ////    {
        ////        for (int month = 1; month <= 6; month++)
        ////        {
        ////            int globalMonth = (halfYear - 1) * 6 + month;
        ////            var firstDayOfMonth = new DateTime(year, globalMonth, 1);
        ////            int daysInMonth = DateTime.DaysInMonth(year, globalMonth);

        ////            ws.Cell(currentRow, currentColumn).Value = firstDayOfMonth.ToString("MMMM");
        ////            ws.Cell(currentRow, currentColumn).Style.Font.Bold = true;
        ////            ws.Range(currentRow, currentColumn, currentRow, currentColumn + 1).Merge();

        ////            currentRow++;

        ////            for (int day = 1; day <= daysInMonth; day++)
        ////            {
        ////                var date = new DateTime(year, globalMonth, day);
        ////                string dayOfWeek = date.ToString("ddd");
        ////                bool isWeekend = date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;

        ////                var eventText = GetEventText(calendars, date);

        ////                ws.Cell(currentRow, currentColumn).Value = day;
        ////                ws.Cell(currentRow, currentColumn + 1).Value = string.IsNullOrWhiteSpace(eventText) ? dayOfWeek : eventText;

        ////                if (isWeekend)
        ////                {
        ////                    ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.LightGray;
        ////                    ws.Cell(currentRow, currentColumn + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
        ////                }

        ////                currentRow++;
        ////            }

        ////            currentRow = startRow;
        ////            currentColumn += 2;
        ////        }

        ////        startRow += 31 + 1;
        ////        currentRow = startRow;
        ////        currentColumn = 1;
        ////    }

        ////    ws.Columns().AdjustToContents();

        ////    using var ms = new MemoryStream();
        ////    workbook.SaveAs(ms);
        ////    return ms.ToArray();
        ////}


        // 3. try
        private byte[] GenerateExcelFile(int year, List<Calendar> calendars)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.AddWorksheet($"{year} Calendar");

            int startRow = 1;
            int currentRow = startRow;
            int currentColumn = 1;
            int dayColumnWidth = 5;
            int eventColumnWidth = 25;

            for (int halfYear = 1; halfYear <= 2; halfYear++)
            {
                for (int month = 1; month <= 6; month++)
                {
                    int globalMonth = (halfYear - 1) * 6 + month;
                    var firstDayOfMonth = new DateTime(year, globalMonth, 1);
                    int daysInMonth = DateTime.DaysInMonth(year, globalMonth);

                    ws.Cell(currentRow, currentColumn).Value = firstDayOfMonth.ToString("MMMM");
                    ws.Cell(currentRow, currentColumn).Style.Font.Bold = true;
                    ws.Range(currentRow, currentColumn, currentRow, currentColumn + 1).Merge();

                    currentRow++;

                    for (int day = 1; day <= daysInMonth; day++)
                    {
                        var date = new DateTime(year, globalMonth, day);
                        string dayOfWeek = date.ToString("ddd");
                        bool isWeekend = date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;

                        var eventText = GetEventText(calendars, date);

                        ws.Cell(currentRow, currentColumn).Value = day;
                        ws.Cell(currentRow, currentColumn + 1).Value = string.IsNullOrWhiteSpace(eventText) ? dayOfWeek : eventText;

                        if (isWeekend)
                        {
                            ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.LightGray;
                            ws.Cell(currentRow, currentColumn + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        }

                        currentRow++;
                    }

                    ws.Column(currentColumn).Width = dayColumnWidth;
                    ws.Column(currentColumn + 1).Width = eventColumnWidth;

                    currentRow = startRow;
                    currentColumn += 2;
                }

                startRow += 31 + 1;
                currentRow = startRow;
                currentColumn = 1;
            }

            // Set print area and page setup options
            var printArea = ws.RangeUsed();
            ws.NamedRanges.Add("CustomPrintArea", printArea);
            ws.PageSetup.PrintAreas.Clear();
            ws.PageSetup.PrintAreas.Add("CustomPrintArea");
            ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            ws.PageSetup.FitToPages(1, 2); // Fit to 1 page wide and 2 pages tall

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        private string GetEventText(List<Calendar> calendars, DateTime date)
        {
            foreach (var calendar in calendars)
            {
                foreach (var e in calendar.Events)
                {
                    if (e.Start.Date == date.Date)
                    {
                        string startTime = e.Start.Value.TimeOfDay != TimeSpan.Zero ? e.Start.Value.ToString("HH:mm") : "";
                        string eventName = e.Summary;

                        foreach (var replacement in _calendarSettings.Value.EventNameReplacements)
                        {
                            eventName = eventName.Replace(replacement.Key, replacement.Value);
                        }

                        return !string.IsNullOrEmpty(startTime) ? $"{startTime} {eventName}" : eventName;
                    }
                }
            }
            return null;
        }
    }
}