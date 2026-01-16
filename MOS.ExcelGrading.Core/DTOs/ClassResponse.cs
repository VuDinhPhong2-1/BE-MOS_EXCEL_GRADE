namespace MOS.ExcelGrading.Core.DTOs
{
    public class ClassResponse
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Code { get; set; } = string.Empty;
        public string SchoolId { get; set; } = string.Empty;
        public string? SchoolName { get; set; }
        public string OwnerId { get; set; } = string.Empty;
        public string? OwnerName { get; set; }
        public string? Description { get; set; }
        public int? MaxStudents { get; set; }
        public int CurrentStudents { get; set; }
        public string? AcademicYear { get; set; }
        public string? Grade { get; set; }
        public List<string> StudentIds { get; set; } = new List<string>();
        public DateTime CreatedAt { get; set; }
        public bool IsActive { get; set; }
    }
}
