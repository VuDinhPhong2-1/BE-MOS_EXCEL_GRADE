namespace MOS.ExcelGrading.Core.DTOs
{
    public class AuthResponse
    {
        public string AccessToken { get; set; } = string.Empty;
        public string RefreshToken { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string UserId { get; set; } = string.Empty;
        public string Role { get; set; } = string.Empty;
        public List<string> Permissions { get; set; } = new List<string>();
        public string? FullName { get; set; }
        public string? Avatar { get; set; }
        public string? TeacherApprovalStatus { get; set; }
        public DateTime? TeacherApprovalRequestedAt { get; set; }
        public DateTime? TeacherApprovalReviewedAt { get; set; }
        public string? TeacherApprovalReviewedBy { get; set; }
        public string? TeacherApprovalNote { get; set; }
        public bool HasPassword { get; set; }
        public bool HasGoogleLinked { get; set; }
    }
}