using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class ExamSession
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("publicationId")]
        public string PublicationId { get; set; } = string.Empty;

        [BsonElement("publicationToken")]
        public string PublicationToken { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("studentId")]
        public string? StudentId { get; set; }

        [BsonElement("studentName")]
        public string StudentName { get; set; } = string.Empty;

        [BsonElement("currentProjectIndex")]
        public int CurrentProjectIndex { get; set; } = 0;

        [BsonElement("currentProjectStatus")]
        public string CurrentProjectStatus { get; set; } = ExamSessionProjectStatuses.NotStarted;

        [BsonElement("status")]
        public string Status { get; set; } = ExamSessionStatuses.InProgress;

        [BsonElement("isAdvancing")]
        public bool IsAdvancing { get; set; } = false;

        [BsonElement("advanceStartedAt")]
        public DateTime? AdvanceStartedAt { get; set; }


        [BsonElement("projectAttempts")]
        public List<ExamSessionProjectAttempt> ProjectAttempts { get; set; } = new();

        [BsonElement("completedProjectCount")]
        public int CompletedProjectCount { get; set; } = 0;

        [BsonElement("aggregateScore")]
        public double? AggregateScore { get; set; }

        [BsonElement("startedAt")]
        public DateTime StartedAt { get; set; } = DateTime.UtcNow;

        [BsonElement("completedAt")]
        public DateTime? CompletedAt { get; set; }

        [BsonElement("lastError")]
        public string? LastError { get; set; }

        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonElement("updatedAt")]
        public DateTime? UpdatedAt { get; set; }

        [BsonElement("version")]
        public int Version { get; set; } = 1;
    }

    public class ExamSessionProjectAttempt
    {
        [BsonElement("projectCode")]
        public string ProjectCode { get; set; } = string.Empty;

        [BsonElement("subject")]
        public string Subject { get; set; } = AssignmentFileSubjects.Excel;

        [BsonElement("templateFileName")]
        public string? TemplateFileName { get; set; }

        [BsonElement("gradingApiEndpoint")]
        public string GradingApiEndpoint { get; set; } = string.Empty;

        [BsonElement("workingFilePath")]
        public string? WorkingFilePath { get; set; }

        [BsonElement("status")]
        public string Status { get; set; } = ExamSessionProjectStatuses.NotStarted;

        [BsonElement("startedAt")]
        public DateTime? StartedAt { get; set; }

        [BsonElement("submittedAt")]
        public DateTime? SubmittedAt { get; set; }

        [BsonElement("gradedAt")]
        public DateTime? GradedAt { get; set; }

        [BsonElement("score")]
        public double? Score { get; set; }

        [BsonElement("maxScore")]
        public double? MaxScore { get; set; }

        [BsonElement("feedback")]
        public string? Feedback { get; set; }

        [BsonElement("attemptNo")]
        public int AttemptNo { get; set; } = 1;
    }

    public static class ExamSessionStatuses
    {
        public const string InProgress = "in_progress";
        public const string Completed = "completed";
        public const string Cancelled = "cancelled";
    }

    public static class ExamSessionProjectStatuses
    {
        public const string NotStarted = "not_started";
        public const string InProgress = "in_progress";
        public const string Submitted = "submitted";
        public const string Graded = "graded";
        public const string Error = "error";
    }
}
