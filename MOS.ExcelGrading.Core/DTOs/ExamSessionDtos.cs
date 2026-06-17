namespace MOS.ExcelGrading.Core.DTOs
{
    public class StartExamSessionRequest
    {
        public string? StudentId { get; set; }
        public string StudentName { get; set; } = string.Empty;
    }

    public class RestartCurrentProjectResponse
    {
        public ExamSessionStateDto State { get; set; } = new();
        public ExamSessionProjectBootstrapDto Bootstrap { get; set; } = new();
    }

    public class ExamSessionStateDto
    {
        public string SessionId { get; set; } = string.Empty;
        public string PublicationId { get; set; } = string.Empty;
        public string PublicationName { get; set; } = string.Empty;

        public int CurrentProjectIndex { get; set; }
        public int CurrentProjectNumber { get; set; }
        public int TotalProjectCount { get; set; }
        public int CompletedProjectCount { get; set; }

        public string Status { get; set; } = string.Empty;
        public string CurrentProjectStatus { get; set; } = string.Empty;
        public bool IsAdvancing { get; set; }
        public DateTime? AdvanceStartedAt { get; set; }
        public DateTime? CompletedAt { get; set; }

        public ExamSessionProjectBootstrapDto? CurrentProject { get; set; }
        public ExamSessionProjectBootstrapDto? NextProject { get; set; }

        public double? AggregateScore { get; set; }
        public string? LastError { get; set; }
    }

    public class ExamSessionProjectBootstrapDto
    {
        public int Order { get; set; }
        public string ProjectCode { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string? TemplateFileName { get; set; }
        public string GradingApiEndpoint { get; set; } = string.Empty;
        public List<ExamPublicationTaskSnapshotItemDto> TaskSnapshot { get; set; } = new();
        public ExamPublicationModeRulesDto? ModeRules { get; set; }
    }

    public class ScoreUploadRequest
    {
        public string? WorkingFilePath { get; set; }
        public double? Score { get; set; }
        public double? MaxScore { get; set; }
        public string? Feedback { get; set; }
        public bool IsSuccess { get; set; } = true;
        public string? ErrorMessage { get; set; }
    }

    public class AdvanceExamSessionResponse
    {
        public bool IsCompleted { get; set; }
        public ExamSessionStateDto State { get; set; } = new();
    }
}
