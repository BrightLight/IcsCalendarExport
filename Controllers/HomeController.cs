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

            var excelFile = await GenerateExcelFile(input.Year, calendars, input.LogoUrl);
            var output = new CalendarOutputModel
            {
                FileName = $"Calendar_{input.Year}.xlsx",
                FileContent = excelFile
            };

            return View("Download", output);
        }

        private async Task<byte[]> GenerateExcelFile(int year, List<Calendar> calendars, string? logoUrl = null)
        {
            using var workbook = new XLWorkbook();
            var wsFirstHalf = workbook.AddWorksheet($"{year} Calendar - H1");
            var wsSecondHalf = workbook.AddWorksheet($"{year} Calendar - H2");

            int dayColumnWidth = 5;
            int eventColumnWidth = 25;
            double dayRowHeight = 22; // Default row height
            double headerRowHeight = 50; // Specific row height for headernt dayRowHeight = 22;
            double eventFontSize = 12; // Default font size

            async Task<byte[]> DownloadImageAsync(string url)
            {
                using var httpClient = new HttpClient();
                return await httpClient.GetByteArrayAsync(url);
            }

            async Task AddHeaderAsync(IXLWorksheet ws, string? url)
            {
                // Merge cells for the header text
                ws.Range("A1:D1").Merge().Value = $"Kalender {year}";
                ws.Range("A1:D1").Style.Font.Bold = true;
                ws.Range("A1:D1").Style.Font.FontSize = 30;
                ws.Range("A1:D1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                // Set the height for the header row
                ws.Row(1).Height = headerRowHeight;
                
                if (!string.IsNullOrWhiteSpace(url))
                {
                    // Download the logo image
                    var logoData = await DownloadImageAsync(url);
                    using var stream = new MemoryStream(logoData);

                    // Add the logo to the top right of the worksheet
                    var logo = ws.AddPicture(stream)
                                  .MoveTo(ws.Cell("G1"))
                                  .Scale(0.5); // Adjust the scale as needed
                }
            }

            async Task FillWorksheetAsync(IXLWorksheet ws, int startMonth, int endMonth, string? url)
            {
                await AddHeaderAsync(ws, url);

                // Set the default row height for the worksheet
                ws.RowHeight = dayRowHeight;

                int startRow = 2; // Adjust startRow to leave space for the header
                int currentRow = startRow;
                int currentColumn = 1;

                for (int month = startMonth; month <= endMonth; month++)
                {
                    var firstDayOfMonth = new DateTime(year, month, 1);
                    int daysInMonth = DateTime.DaysInMonth(year, month);

                    ws.Cell(currentRow, currentColumn).Value = firstDayOfMonth.ToString("MMMM");
                    ws.Cell(currentRow, currentColumn).Style.Font.Bold = true;
                    ws.Range(currentRow, currentColumn, currentRow, currentColumn + 1).Merge();

                    currentRow++;

                    for (int day = 1; day <= daysInMonth; day++)
                    {
                        var date = new DateTime(year, month, day);
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

                        // Set the font size for the day and event text
                        ws.Cell(currentRow, currentColumn).Style.Font.FontSize = eventFontSize;
                        ws.Cell(currentRow, currentColumn + 1).Style.Font.FontSize = eventFontSize;

                        // Set the row height for each day row
                        ws.Row(currentRow).Height = dayRowHeight;

                        currentRow++;
                    }

                    ws.Column(currentColumn).Width = dayColumnWidth;
                    ws.Column(currentColumn + 1).Width = eventColumnWidth;

                    currentRow = startRow;
                    currentColumn += 2;
                }
            }

            // Fill the first worksheet with the first half of the year
            await FillWorksheetAsync(wsFirstHalf, 1, 6, logoUrl);

            // Fill the second worksheet with the second half of the year
            await FillWorksheetAsync(wsSecondHalf, 7, 12, logoUrl);

            // Set print area and page setup options for the first half
            var printAreaFirstHalf = wsFirstHalf.RangeUsed();
            wsFirstHalf.NamedRanges.Add("CustomPrintAreaH1", printAreaFirstHalf);
            wsFirstHalf.PageSetup.PrintAreas.Clear();
            wsFirstHalf.PageSetup.PrintAreas.Add("CustomPrintAreaH1");
            wsFirstHalf.PageSetup.PaperSize = XLPaperSize.A4Paper;
            wsFirstHalf.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            wsFirstHalf.PageSetup.FitToPages(1, 2); // Fit to 1 page wide and 2 pages tall

            // Set print area and page setup options for the second half
            var printAreaSecondHalf = wsSecondHalf.RangeUsed();
            wsSecondHalf.NamedRanges.Add("CustomPrintAreaH2", printAreaSecondHalf);
            wsSecondHalf.PageSetup.PrintAreas.Clear();
            wsSecondHalf.PageSetup.PrintAreas.Add("CustomPrintAreaH2");
            wsSecondHalf.PageSetup.PaperSize = XLPaperSize.A4Paper;
            wsSecondHalf.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            wsSecondHalf.PageSetup.FitToPages(1, 2); // Fit to 1 page wide and 2 pages tall

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }

        private string? GetEventText(List<Calendar> calendars, DateTime date)
        {
            // Define your local time zone. For example, using Central European Time (CET)
            TimeZoneInfo localTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Central European Standard Time");

            foreach (var calendar in calendars)
            {
                foreach (var e in calendar.Events)
                {
                    DateTime localStartTime;
                    // Convert the event start time from UTC to the local time zone
                    if (e.Start.IsUtc)
                    {
                        // Ensure the event start time is treated as UTC
                        DateTime utcStartTime = DateTime.SpecifyKind(e.Start.Value, DateTimeKind.Utc);

                        // Convert the event start time from UTC to the local time zone
                        localStartTime = TimeZoneInfo.ConvertTimeFromUtc(utcStartTime, localTimeZone);
                    }
                    else
                    {
                        // If the event start time is already in the local time zone, use it directly
                        localStartTime = e.Start.Value;
                    }

                    if (localStartTime.Date == date.Date)
                    {
                        string startTime = localStartTime.TimeOfDay != TimeSpan.Zero ? localStartTime.ToString("HH:mm") : "";
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