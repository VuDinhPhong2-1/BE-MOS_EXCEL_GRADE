namespace MOS.ExcelGrading.Core.DTOs
{
    public class TeacherApprovalDecisionRequest
    {
        public string Decision { get; set; } = string.Empty;
        public string? Note { get; set; }
    }

    public class TeacherApprovalResponse
    {
        public string UserId { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string? FullName { get; set; }
        public string Role { get; set; } = string.Empty;
        public List<string> Permissions { get; set; } = new();
        public bool IsActive { get; set; }
        public string? TeacherApprovalStatus { get; set; }
        public DateTime? TeacherApprovalRequestedAt { get; set; }
        public DateTime? TeacherApprovalReviewedAt { get; set; }
        public string? TeacherApprovalReviewedBy { get; set; }
        public string? TeacherApprovalNote { get; set; }
    }
}