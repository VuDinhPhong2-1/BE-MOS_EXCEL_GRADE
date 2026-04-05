namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateGradingTestBugNoteRequest
    {
        public string ProjectCode { get; set; } = string.Empty;
        public string? ProjectDisplayName { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Severity { get; set; } = "medium";
        public GradingTestBugNoteScoreSummaryDto? ScoreSummary { get; set; }
        public string? GradingError { get; set; }
    }

    public class GradingTestBugNoteResponse
    {
        public string Id { get; set; } = string.Empty;
        public string ProjectCode { get; set; } = string.Empty;
        public string ProjectDisplayName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Severity { get; set; } = "medium";
        public GradingTestBugNoteScoreSummaryDto? ScoreSummary { get; set; }
        public string? GradingError { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    public class GradingTestBugNoteScoreSummaryDto
    {
        public double TotalScore { get; set; }
        public double MaxScore { get; set; }
        public double Percentage { get; set; }
        public string Status { get; set; } = string.Empty;
    }
}
