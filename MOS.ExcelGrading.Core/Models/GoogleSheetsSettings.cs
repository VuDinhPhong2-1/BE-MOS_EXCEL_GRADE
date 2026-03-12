namespace MOS.ExcelGrading.Core.Models
{
    public class GoogleSheetsSettings
    {
        public bool Enabled { get; set; }
        public string ApplicationName { get; set; } = "MOS Grader";
        public string? ServiceAccountJson { get; set; }
        public string? ServiceAccountJsonPath { get; set; }
        public string? DefaultSpreadsheetId { get; set; }
    }
}
