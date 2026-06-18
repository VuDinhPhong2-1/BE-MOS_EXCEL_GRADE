namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateExamPublicationRequest
    {
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string? ClassId { get; set; }
        public List<string> StudentIds { get; set; } = new();
        public List<string> AssignmentIds { get; set; } = new();
        public string? Mode { get; set; }
        public DateTime? StartsAt { get; set; }
        public DateTime? EndsAt { get; set; }
        public int? DurationMinutes { get; set; }
        public List<CreateExamPublicationProjectRequest> ProjectSequence { get; set; } = new();

        // Backward-compatible single-project payload fields.
        public string? ProjectCode { get; set; }
        public string? Subject { get; set; }
        public string? TemplateFileName { get; set; }
        public string? GradingApiEndpoint { get; set; }
        public List<ExamPublicationTaskSnapshotItemDto> TaskSnapshot { get; set; } = new();
        public ExamPublicationModeRulesDto? ModeRules { get; set; }
    }

    public class CreateExamPublicationProjectRequest
    {
        public int? Order { get; set; }
        public string? SourceAssignmentId { get; set; }
        public string? ProjectCode { get; set; }
        public string? Subject { get; set; }
        public string? TemplateFileName { get; set; }
        public string GradingApiEndpoint { get; set; } = string.Empty;
        public List<ExamPublicationTaskSnapshotItemDto> TaskSnapshot { get; set; } = new();
        public ExamPublicationModeRulesDto? ModeRules { get; set; }
    }

    public class ExamPublicationTaskSnapshotItemDto
    {
        public string? TaskId { get; set; }
        public string? TaskName { get; set; }
        public double? MaxScore { get; set; }
        public string? Instructions { get; set; }
    }

    public class ExamPublicationModeRulesDto
    {
        public string? Mode { get; set; }
        public bool? ShowFeedback { get; set; }
        public bool? AllowRestart { get; set; }
        public bool? AllowNextProject { get; set; }
    }

    public class PublicExamStudentDto
    {
        public string Id { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
    }

    public class PublicExamPublicationInfoDto
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public DateTime? StartsAt { get; set; }
        public DateTime? EndsAt { get; set; }
        public int? DurationMinutes { get; set; }
        public List<string> StudentIds { get; set; } = new();
        public List<PublicExamStudentDto> Students { get; set; } = new();
        public int ProjectCount { get; set; }
    }
}
