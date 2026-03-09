using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateScheduleRequest
    {
        public string? SchoolId { get; set; }
        public string? ClassId { get; set; }

        [StringLength(150)]
        public string? ClassName { get; set; }

        [Required(ErrorMessage = "Môn học là bắt buộc")]
        [StringLength(120)]
        public string Subject { get; set; } = string.Empty;

        [StringLength(100)]
        public string? RoomName { get; set; }

        [StringLength(60)]
        public string? PeriodLabel { get; set; }

        [Required(ErrorMessage = "Ngày học là bắt buộc")]
        public DateTime Date { get; set; }

        [Required(ErrorMessage = "Giờ bắt đầu là bắt buộc")]
        [RegularExpression("^([01]\\d|2[0-3]):[0-5]\\d$", ErrorMessage = "Giờ bắt đầu phải theo định dạng HH:mm")]
        public string StartTime { get; set; } = string.Empty;

        [Required(ErrorMessage = "Giờ kết thúc là bắt buộc")]
        [RegularExpression("^([01]\\d|2[0-3]):[0-5]\\d$", ErrorMessage = "Giờ kết thúc phải theo định dạng HH:mm")]
        public string EndTime { get; set; } = string.Empty;

        [StringLength(500)]
        public string? Notes { get; set; }
    }

    public class UpdateScheduleRequest
    {
        public string? SchoolId { get; set; }
        public string? ClassId { get; set; }

        [StringLength(150)]
        public string? ClassName { get; set; }

        [Required(ErrorMessage = "Môn học là bắt buộc")]
        [StringLength(120)]
        public string Subject { get; set; } = string.Empty;

        [StringLength(100)]
        public string? RoomName { get; set; }

        [StringLength(60)]
        public string? PeriodLabel { get; set; }

        [Required(ErrorMessage = "Ngày học là bắt buộc")]
        public DateTime Date { get; set; }

        [Required(ErrorMessage = "Giờ bắt đầu là bắt buộc")]
        [RegularExpression("^([01]\\d|2[0-3]):[0-5]\\d$", ErrorMessage = "Giờ bắt đầu phải theo định dạng HH:mm")]
        public string StartTime { get; set; } = string.Empty;

        [Required(ErrorMessage = "Giờ kết thúc là bắt buộc")]
        [RegularExpression("^([01]\\d|2[0-3]):[0-5]\\d$", ErrorMessage = "Giờ kết thúc phải theo định dạng HH:mm")]
        public string EndTime { get; set; } = string.Empty;

        public bool? IsActive { get; set; }

        [StringLength(500)]
        public string? Notes { get; set; }
    }

    public class ScheduleResponse
    {
        public string Id { get; set; } = string.Empty;
        public string OwnerId { get; set; } = string.Empty;
        public string? SchoolId { get; set; }
        public string? ClassId { get; set; }
        public string ClassName { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string? RoomName { get; set; }
        public string? PeriodLabel { get; set; }
        public DateTime Date { get; set; }
        public int DayOfWeek { get; set; }
        public string StartTime { get; set; } = string.Empty;
        public string EndTime { get; set; } = string.Empty;
        public string? Notes { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }
}
