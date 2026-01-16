namespace MOS.ExcelGrading.Core.DTOs
{
    public class BulkImportStudentRequest
    {
        public List<StudentImportItem> Students { get; set; } = new();
        public string? ClassId { get; set; }
    }

    public class StudentImportItem
    {
        public string MiddleName { get; set; } = string.Empty;  // MiddleName
        public string FirstName { get; set; } = string.Empty;    // FirstName
    }

    public class BulkImportResult
    {
        public int TotalCount { get; set; }
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public List<string> Errors { get; set; } = new();
        public List<StudentResponse> ImportedStudents { get; set; } = new();
    }
}
