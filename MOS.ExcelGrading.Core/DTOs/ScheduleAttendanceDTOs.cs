using System.ComponentModel.DataAnnotations;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class SaveScheduleAttendanceRequest
    {
        [Required(ErrorMessage = "Danh sách điểm danh không được để trống")]
        public List<SaveScheduleAttendanceItem> Items { get; set; } = new();
    }

    public class SaveScheduleAttendanceItem
    {
        [Required(ErrorMessage = "Thiếu mã học sinh")]
        public string StudentId { get; set; } = string.Empty;

        [Required(ErrorMessage = "Thiếu trạng thái điểm danh")]
        [RegularExpression("^(Present|Absent)$", ErrorMessage = "Trạng thái điểm danh không hợp lệ")]
        public string Status { get; set; } = AttendanceStatus.Present;

        [StringLength(500)]
        public string? Note { get; set; }
    }

    public class ScheduleAttendanceStudentResponse
    {
        public string StudentId { get; set; } = string.Empty;
        public string MiddleName { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string StudentStatus { get; set; } = string.Empty;
        public string AttendanceStatus { get; set; } = MOS.ExcelGrading.Core.Models.AttendanceStatus.Present;
        public string? Note { get; set; }
        public DateTime? MarkedAt { get; set; }
    }

    public class ScheduleAttendanceResponse
    {
        public string ScheduleId { get; set; } = string.Empty;
        public string? ClassId { get; set; }
        public string ClassName { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public string StartTime { get; set; } = string.Empty;
        public string EndTime { get; set; } = string.Empty;
        public string? RoomName { get; set; }
        public List<ScheduleAttendanceStudentResponse> Students { get; set; } = new();
        public int PresentCount { get; set; }
        public int AbsentCount { get; set; }
    }
}
