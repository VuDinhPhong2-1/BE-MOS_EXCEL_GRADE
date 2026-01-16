namespace MOS.ExcelGrading.Core.DTOs
{
    public class UpdateClassRequest
    {
        public string? Name { get; set; }
        public string? Code { get; set; }
        public string? Description { get; set; }
        public int? MaxStudents { get; set; }
        public string? AcademicYear { get; set; }
        public string? Grade { get; set; }
        public string? TeacherId { get; set; }
        public bool? IsActive { get; set; }
    }
}
