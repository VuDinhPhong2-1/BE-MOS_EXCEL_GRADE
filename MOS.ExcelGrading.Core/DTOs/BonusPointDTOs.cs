// MOS.ExcelGrading.Core/DTOs/BonusPointDTOs.cs
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateBonusPointRequest
    {
        [Required(ErrorMessage = "Mã học sinh là bắt buộc")]
        public string StudentId { get; set; } = string.Empty;

        [Required(ErrorMessage = "Mã lớp là bắt buộc")]
        public string ClassId { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ngày là bắt buộc")]
        public DateTime Date { get; set; }

        public double Points { get; set; }

        public string Category { get; set; } = "participation";

        public string? Reason { get; set; }

        public string? ScheduleId { get; set; }
    }

    public class UpdateBonusPointRequest
    {
        public DateTime? Date { get; set; }
        public double? Points { get; set; }
        public string? Category { get; set; }
        public string? Reason { get; set; }
        public string? ScheduleId { get; set; }
    }

    public class BulkBonusPointRequest
    {
        [Required(ErrorMessage = "Mã lớp là bắt buộc")]
        public string ClassId { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ngày là bắt buộc")]
        public DateTime Date { get; set; }

        public string? ScheduleId { get; set; }

        public List<StudentBonusItem> Items { get; set; } = new();
    }

    public class StudentBonusItem
    {
        [Required(ErrorMessage = "Mã học sinh là bắt buộc")]
        public string StudentId { get; set; } = string.Empty;

        public double Points { get; set; }

        public string Category { get; set; } = "participation";

        public string? Reason { get; set; }
    }

    public class BonusPointResponse
    {
        public string Id { get; set; } = string.Empty;
        public string StudentId { get; set; } = string.Empty;
        public string StudentFirstName { get; set; } = string.Empty;
        public string StudentMiddleName { get; set; } = string.Empty;
        public string StudentFullName { get; set; } = string.Empty;
        public string ClassId { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public double Points { get; set; }
        public string Category { get; set; } = string.Empty;
        public string? Reason { get; set; }
        public string? ScheduleId { get; set; }
        public DateTime CreatedAt { get; set; }
        public string? CreatedBy { get; set; }
        public string? CreatedByName { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }

    public class StudentBonusPointSummary
    {
        public string StudentId { get; set; } = string.Empty;
        public string StudentFirstName { get; set; } = string.Empty;
        public string StudentMiddleName { get; set; } = string.Empty;
        public string StudentFullName { get; set; } = string.Empty;
        public double TotalBonusPoints { get; set; }
        public List<BonusPointResponse> Details { get; set; } = new();
    }

    public class ClassBonusSummaryResponse
    {
        public string ClassId { get; set; } = string.Empty;
        public List<StudentBonusPointSummary> Students { get; set; } = new();
    }
}
