using Microsoft.AspNetCore.Http;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class ImportStudentRequest
    {
        public IFormFile ExcelFile { get; set; } = null!;
        public string? TeacherId { get; set; }
        public string? TeacherName { get; set; }
    }

    public class ImportStudentResult
    {
        public int TotalRows { get; set; }
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public List<string> Errors { get; set; } = new();
        public List<StudentResponse> ImportedStudents { get; set; } = new();
    }

    public class ExcelStudentRow
    {
        public int RowNumber { get; set; }
        public string MiddleName { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
    }
}
