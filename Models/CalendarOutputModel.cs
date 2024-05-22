namespace IcsCalendar2Excel.Models
{
    // Models/CalendarOutputModel.cs
    public class CalendarOutputModel
    {
        /// <summary>
        /// Gets or sets the name of the file to download.
        /// </summary>
        public required string FileName { get; init; }
        
        /// <summary>
        /// Gets or sets the content of the file to download.
        /// </summary>
        public required byte[] FileContent { get; init; }
    }

}
