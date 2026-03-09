namespace MOS.ExcelGrading.Core.Models
{
    public class ScheduleReportBundle
    {
        public StartLessonReport StartLesson { get; set; } = new();
        public ProfessionalReport Professional { get; set; } = new();
        public EndLessonReport EndLesson { get; set; } = new();
    }

    public class StartLessonReport
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

    public class ProfessionalReport
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

    public class EndLessonReport
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
}
