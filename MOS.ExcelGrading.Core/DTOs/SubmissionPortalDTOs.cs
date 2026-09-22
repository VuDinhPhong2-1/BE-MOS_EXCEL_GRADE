namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateSubmissionPortalRequest
    {
        public string Title { get; set; } = string.Empty;
        public string? Description { get; set; }
        public List<string> ClassIds { get; set; } = new();
        public List<string> AssignmentIds { get; set; } = new();
        public DateTime? StartsAt { get; set; }
        public DateTime? EndsAt { get; set; }
        public int MaxSubmissionsPerStudent { get; set; }
        public string ScoringPolicy { get; set; } = "BestScore";
        public bool ShowLeaderboard { get; set; } = true;
        public bool ShowDetailedFeedback { get; set; } = true;
    }

    public class UpdateSubmissionPortalRequest : CreateSubmissionPortalRequest
    {
        public bool IsActive { get; set; } = true;
    }

    public class SubmissionPortalResponse
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string PublicToken { get; set; } = string.Empty;
        public List<string> ClassIds { get; set; } = new();
        public List<string> AssignmentIds { get; set; } = new();
        public DateTime? StartsAt { get; set; }
        public DateTime? EndsAt { get; set; }
        public int MaxSubmissionsPerStudent { get; set; }
        public string ScoringPolicy { get; set; } = "BestScore";
        public bool ShowLeaderboard { get; set; }
        public bool ShowDetailedFeedback { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public int UnreadAlertCount { get; set; }
    }

    public class PublicPortalInfoResponse
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string? Description { get; set; }
        public DateTime? StartsAt { get; set; }
        public DateTime? EndsAt { get; set; }
        public int MaxSubmissionsPerStudent { get; set; }
        public string ScoringPolicy { get; set; } = "BestScore";
        public bool ShowLeaderboard { get; set; }
        public bool ShowDetailedFeedback { get; set; }
        public List<PublicPortalClassDto> Classes { get; set; } = new();
        public List<PublicPortalAssignmentDto> Assignments { get; set; } = new();
    }

    public class PublicPortalClassDto
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    public class PublicPortalStudentDto
    {
        public string Id { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string ClassId { get; set; } = string.Empty;
    }

    public class PublicPortalAssignmentDto
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string ClassId { get; set; } = string.Empty;
        public double MaxScore { get; set; }
        public string Subject { get; set; } = "excel";
        public string? GradingApiEndpoint { get; set; }
        public bool HasTemplate { get; set; }
        public bool HasInstructions { get; set; }
    }

    public class PublicPortalSubmitResult
    {
        public double? ScoreValue { get; set; }
        public double MaxScore { get; set; }
        public string? Feedback { get; set; }
        public List<string> AutoGradingErrors { get; set; } = new();
        public List<AutoGradingTaskResultRequest> AutoGradingTaskResults { get; set; } = new();
        public DateTime SubmittedAt { get; set; }
        public int? Rank { get; set; }
        public List<string> Alerts { get; set; } = new();
    }

    public class SubmissionLeaderboardItem
    {
        public int Rank { get; set; }
        public string StudentId { get; set; } = string.Empty;
        public string StudentName { get; set; } = string.Empty;
        public string ClassId { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public string? AssignmentId { get; set; }
        public string? AssignmentName { get; set; }
        public double ScoreValue { get; set; }
        public double MaxScore { get; set; }
        public DateTime? GradedAt { get; set; }
        public int SubmissionCount { get; set; }
    }

    public class SubmissionLogResponse
    {
        public string Id { get; set; } = string.Empty;
        public string StudentId { get; set; } = string.Empty;
        public string StudentName { get; set; } = string.Empty;
        public string ClassId { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public string AssignmentId { get; set; } = string.Empty;
        public string AssignmentName { get; set; } = string.Empty;
        public double? ScoreValue { get; set; }
        public double MaxScore { get; set; }
        public string? IpAddress { get; set; }
        public string? FileHash { get; set; }
        public string? FileName { get; set; }
        public long? FileSizeBytes { get; set; }
        public List<string> Alerts { get; set; } = new();
        public DateTime SubmittedAt { get; set; }
    }

    public class SubmissionAlertResponse
    {
        public string Id { get; set; } = string.Empty;
        public string AlertType { get; set; } = string.Empty;
        public string Severity { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public List<string> InvolvedStudentIds { get; set; } = new();
        public List<string> InvolvedSubmissionLogIds { get; set; } = new();
        public bool IsRead { get; set; }
        public bool IsDismissed { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}