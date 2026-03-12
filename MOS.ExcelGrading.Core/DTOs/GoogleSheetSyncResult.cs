namespace MOS.ExcelGrading.Core.DTOs
{
    public class GoogleSheetSyncResult
    {
        public string SpreadsheetId { get; set; } = string.Empty;
        public string WorksheetName { get; set; } = string.Empty;
        public int ColumnIndex { get; set; }
        public string ColumnLetter { get; set; } = string.Empty;
        public int ClassificationColumnIndex { get; set; }
        public string ClassificationColumnLetter { get; set; } = string.Empty;
        public int NotesColumnIndex { get; set; }
        public string NotesColumnLetter { get; set; } = string.Empty;
        public int MatchedStudentCount { get; set; }
        public int TotalStudentCount { get; set; }
        public int PresentMarkedCount { get; set; }
        public int AbsentClearedCount { get; set; }
        public int UnmatchedSheetRowCount { get; set; }
        public int UnmatchedAttendanceStudentCount { get; set; }
        public int UnmatchedStudentCount { get; set; }
        public bool UsedFallbackNameMatch { get; set; }
    }
}
