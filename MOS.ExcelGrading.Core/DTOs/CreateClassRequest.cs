using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateClassRequest
    {
        [Required]
        [StringLength(100, MinimumLength = 2)]
        public string Name { get; set; } = string.Empty;

        [Required]
        public string SchoolId { get; set; } = string.Empty;

        public string? Description { get; set; }
        public int? MaxStudents { get; set; }
        public string? AcademicYear { get; set; }
        public string? Grade { get; set; }
    }
}
