using System.ComponentModel.DataAnnotations;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class SaveScheduleAttendanceRequest
    {
        [Required(ErrorMessage = "Danh sach diem danh khong duoc de trong")]
        public List<SaveScheduleAttendanceItem> Items { get; set; } = new();

        public ScheduleReportsRequest? Reports { get; set; }
    }

    public class SaveScheduleAttendanceItem
    {
        [Required(ErrorMessage = "Thieu ma hoc sinh")]
        public string StudentId { get; set; } = string.Empty;

        [Required(ErrorMessage = "Thieu trang thai diem danh")]
        [RegularExpression("^(Present|Absent)$", ErrorMessage = "Trang thai diem danh khong hop le")]
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
        public string? SchoolId { get; set; }
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
        public ScheduleReportsResponse Reports { get; set; } = new();
        public ScheduleRoomSessionContextResponse RoomSessionContext { get; set; } = new();
    }

    public class ScheduleReportsRequest
    {
        public StartLessonReportRequest StartLesson { get; set; } = new();
        public ProfessionalReportRequest Professional { get; set; } = new();
        public EndLessonReportRequest EndLesson { get; set; } = new();
    }

    public class StartLessonReportRequest
    {
        [StringLength(200)]
        public string? TeacherName { get; set; }

        [StringLength(200)]
        public string? AssistantName { get; set; }

        [StringLength(120)]
        public string? RoomName { get; set; }

        [StringLength(300)]
        public string? TotalMachines { get; set; }

        [StringLength(500)]
        public string? BrokenMachinesSummary { get; set; }

        [StringLength(300)]
        public string? MissingMachinesForStudents { get; set; }

        [StringLength(120)]
        public string? NetSupportStatus { get; set; }

        [StringLength(120)]
        public string? AudioStatus { get; set; }

        [StringLength(120)]
        public string? CoolingStatus { get; set; }

        [StringLength(120)]
        public string? HygieneStatus { get; set; }
    }

    public class ProfessionalReportRequest
    {
        [StringLength(200)]
        public string? TeacherName { get; set; }

        [StringLength(150)]
        public string? ClassName { get; set; }

        [StringLength(150)]
        public string? SubjectName { get; set; }

        [StringLength(200)]
        public string? TeachingMaterials { get; set; }

        [StringLength(500)]
        public string? TeachingContent { get; set; }

        [StringLength(100)]
        public string? PlannedLessons { get; set; }

        [StringLength(100)]
        public string? TaughtLessons { get; set; }

        [StringLength(100)]
        public string? OngoingPracticeCompletions { get; set; }

        [StringLength(300)]
        public string? GmetrixResultRate { get; set; }
    }

    public class EndLessonReportRequest
    {
        [StringLength(200)]
        public string? TeacherName { get; set; }

        [StringLength(200)]
        public string? AssistantName { get; set; }

        [StringLength(120)]
        public string? RoomName { get; set; }

        [StringLength(300)]
        public string? TotalMachines { get; set; }

        [StringLength(500)]
        public string? ClassStudentCountSummary { get; set; }

        [StringLength(120)]
        public string? StudentMaterialCoverageRate { get; set; }

        [StringLength(500)]
        public string? BrokenMachinesSummary { get; set; }

        [StringLength(120)]
        public string? NetSupportStatus { get; set; }

        [StringLength(120)]
        public string? AudioStatus { get; set; }

        [StringLength(120)]
        public string? CoolingStatus { get; set; }

        [StringLength(120)]
        public string? DevicesPoweredOffStatus { get; set; }

        [StringLength(120)]
        public string? SeatingOrderStatus { get; set; }

        [StringLength(120)]
        public string? RoomHygieneStatus { get; set; }

        [StringLength(120)]
        public string? StudentRuleComplianceStatus { get; set; }

        [StringLength(300)]
        public string? ViolationListSummary { get; set; }
    }

    public class ScheduleReportsResponse
    {
        public StartLessonReportResponse StartLesson { get; set; } = new();
        public ProfessionalReportResponse Professional { get; set; } = new();
        public EndLessonReportResponse EndLesson { get; set; } = new();
    }

    public class StartLessonReportResponse
    {
        public string TeacherName { get; set; } = string.Empty;
        public string AssistantName { get; set; } = string.Empty;
        public string RoomName { get; set; } = string.Empty;
        public string TotalMachines { get; set; } = string.Empty;
        public string BrokenMachinesSummary { get; set; } = string.Empty;
        public string MissingMachinesForStudents { get; set; } = string.Empty;
        public string NetSupportStatus { get; set; } = string.Empty;
        public string AudioStatus { get; set; } = string.Empty;
        public string CoolingStatus { get; set; } = string.Empty;
        public string HygieneStatus { get; set; } = string.Empty;
    }

    public class ProfessionalReportResponse
    {
        public string TeacherName { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public string SubjectName { get; set; } = string.Empty;
        public string TeachingMaterials { get; set; } = string.Empty;
        public string TeachingContent { get; set; } = string.Empty;
        public string PlannedLessons { get; set; } = string.Empty;
        public string TaughtLessons { get; set; } = string.Empty;
        public string OngoingPracticeCompletions { get; set; } = string.Empty;
        public string GmetrixResultRate { get; set; } = string.Empty;
    }

    public class EndLessonReportResponse
    {
        public string TeacherName { get; set; } = string.Empty;
        public string AssistantName { get; set; } = string.Empty;
        public string RoomName { get; set; } = string.Empty;
        public string TotalMachines { get; set; } = string.Empty;
        public string ClassStudentCountSummary { get; set; } = string.Empty;
        public string StudentMaterialCoverageRate { get; set; } = string.Empty;
        public string BrokenMachinesSummary { get; set; } = string.Empty;
        public string NetSupportStatus { get; set; } = string.Empty;
        public string AudioStatus { get; set; } = string.Empty;
        public string CoolingStatus { get; set; } = string.Empty;
        public string DevicesPoweredOffStatus { get; set; } = string.Empty;
        public string SeatingOrderStatus { get; set; } = string.Empty;
        public string RoomHygieneStatus { get; set; } = string.Empty;
        public string StudentRuleComplianceStatus { get; set; } = string.Empty;
        public string ViolationListSummary { get; set; } = string.Empty;
    }

    public class ScheduleRoomSessionContextResponse
    {
        public string SessionLabel { get; set; } = string.Empty;
        public bool IsSharedRoomSession { get; set; }
        public string SharedClassStudentSummary { get; set; } = string.Empty;
        public List<ScheduleRoomClassSummaryResponse> SharedClasses { get; set; } = new();
    }

    public class ScheduleRoomClassSummaryResponse
    {
        public string? ClassId { get; set; }
        public string ClassName { get; set; } = string.Empty;
        public int CurrentStudents { get; set; }
        public int? MaxStudents { get; set; }
    }
}
