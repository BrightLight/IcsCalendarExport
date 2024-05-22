namespace IcsCalendar2Excel.Models
{
    // Models/CalendarInputModel.cs
    public class CalendarInputModel
    {
        /// <summary>
        /// Gets or sets the URLs of the calendars to include in the Excel file.
        /// </summary>
        public List<string>? Urls { get; set; }

        /// <summary>
        /// Gets or sets the year to include in the Excel file.
        /// </summary>
        public int Year { get; set; }

        /// <summary>
        /// Gets or sets the URL of the logo to include in the Excel file.
        /// </summary>
        public string? LogoUrl { get; set; }
    }
}
