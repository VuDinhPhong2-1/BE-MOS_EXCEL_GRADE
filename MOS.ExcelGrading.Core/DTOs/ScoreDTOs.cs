// MOS.ExcelGrading.Core/DTOs/ScoreDTOs.cs
namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateScoreRequest
    {
        public string StudentId { get; set; } = string.Empty;
        public string AssignmentId { get; set; } = string.Empty;
        public string ClassId { get; set; } = string.Empty;
        public double? ScoreValue { get; set; }
        public string? Feedback { get; set; }
    }

    public class UpdateScoreRequest
    {
        public double? ScoreValue { get; set; }
        public string? Feedback { get; set; }
    }

    public class BulkScoreRequest
    {
        public string AssignmentId { get; set; } = string.Empty;
        public string ClassId { get; set; } = string.Empty;
        public List<StudentScoreItem> Scores { get; set; } = new();
    }

    public class StudentScoreItem
    {
        public string StudentId { get; set; } = string.Empty;
        public double? ScoreValue { get; set; }
        public string? Feedback { get; set; }
    }

    public class ScoreResponse
    {
        public string Id { get; set; } = string.Empty;
        public string StudentId { get; set; } = string.Empty;
        public string StudentFirstName { get; set; } = string.Empty;
        public string StudentMiddleName { get; set; } = string.Empty;
        public string StudentFullName { get; set; } = string.Empty;
        public string AssignmentId { get; set; } = string.Empty;
        public string AssignmentName { get; set; } = string.Empty;
        public double? ScoreValue { get; set; }
        public string? Feedback { get; set; }
        public DateTime? GradedAt { get; set; }
        public string? GradedBy { get; set; }
        public string? GradedByName { get; set; }
    }

    public class StudentScoreReportResponse
    {
        public string StudentId { get; set; } = string.Empty;
        public string StudentFullName { get; set; } = string.Empty;
        public List<ScoreDetailResponse> Scores { get; set; } = new();
        public double AverageScore { get; set; }
        public int TotalAssignments { get; set; }
        public int CompletedAssignments { get; set; }
    }

    public class ScoreDetailResponse
    {
        public string AssignmentName { get; set; } = string.Empty;
        public double? ScoreValue { get; set; }
        public double MaxScore { get; set; }
        public string? Feedback { get; set; }
        public DateTime? GradedAt { get; set; }
    }
}
