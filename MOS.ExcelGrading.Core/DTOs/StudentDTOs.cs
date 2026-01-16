// MOS.ExcelGrading.Core/DTOs/StudentDTOs.cs
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateStudentRequest
    {
        [Required]
        [StringLength(100)]
        public string MiddleName { get; set; } = string.Empty;

        [Required]
        [StringLength(100)]
        public string FirstName { get; set; } = string.Empty;

        public string Status { get; set; } = "Active";

        public string? TeacherId { get; set; }
        public string? ClassId { get; set; }
        public string? TeacherName { get; set; }
    }

    public class UpdateStudentRequest
    {
        [StringLength(100)]
        public string? MiddleName { get; set; }

        [StringLength(100)]
        public string? FirstName { get; set; }

        public string? Status { get; set; }

        public string? TeacherId { get; set; }

        public string? TeacherName { get; set; }
        public string? ClassId { get; set; }
    }

    public class StudentResponse
    {
        public string Id { get; set; } = string.Empty;
        public string MiddleName { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string FullName => $"{MiddleName} {FirstName}";
        public string Status { get; set; } = string.Empty;
        public string? TeacherId { get; set; }
        public string? ClassId { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public bool IsActive { get; set; }

    }
}
