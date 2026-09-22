using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class SubmissionPortal
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [Required]
        [StringLength(200)]
        [BsonElement("title")]
        public string Title { get; set; } = string.Empty;

        [StringLength(1000)]
        [BsonElement("description")]
        public string? Description { get; set; }

        [BsonElement("publicToken")]
        public string PublicToken { get; set; } = Guid.NewGuid().ToString("N");

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classIds")]
        public List<string> ClassIds { get; set; } = new();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("assignmentIds")]
        public List<string> AssignmentIds { get; set; } = new();

        [BsonElement("startsAt")]
        public DateTime? StartsAt { get; set; }

        [BsonElement("endsAt")]
        public DateTime? EndsAt { get; set; }

        [BsonElement("maxSubmissionsPerStudent")]
        public int MaxSubmissionsPerStudent { get; set; }

        [BsonElement("scoringPolicy")]
        public string ScoringPolicy { get; set; } = SubmissionPortalScoringPolicies.BestScore;

        [BsonElement("showLeaderboard")]
        public bool ShowLeaderboard { get; set; } = true;

        [BsonElement("showDetailedFeedback")]
        public bool ShowDetailedFeedback { get; set; } = true;

        [BsonElement("isActive")]
        public bool IsActive { get; set; } = true;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("createdBy")]
        public string? CreatedBy { get; set; }

        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonElement("updatedAt")]
        public DateTime? UpdatedAt { get; set; }
    }

    public static class SubmissionPortalScoringPolicies
    {
        public const string BestScore = "BestScore";
        public const string LatestScore = "LatestScore";

        public static bool IsValid(string? value) =>
            string.Equals(value, BestScore, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(value, LatestScore, StringComparison.OrdinalIgnoreCase);

        public static string Normalize(string? value) =>
            string.Equals(value, LatestScore, StringComparison.OrdinalIgnoreCase) ? LatestScore : BestScore;
    }

    [BsonIgnoreExtraElements]
    public class SubmissionLog
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("portalId")]
        public string PortalId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("studentId")]
        public string StudentId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string ClassId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("assignmentId")]
        public string AssignmentId { get; set; } = string.Empty;

        [BsonElement("scoreValue")]
        public double? ScoreValue { get; set; }

        [BsonElement("maxScore")]
        public double MaxScore { get; set; }

        [BsonElement("ipAddress")]
        public string? IpAddress { get; set; }

        [BsonElement("userAgent")]
        public string? UserAgent { get; set; }

        [BsonElement("fileHash")]
        public string? FileHash { get; set; }

        [BsonElement("fileName")]
        public string? FileName { get; set; }

        [BsonElement("fileSizeBytes")]
        public long? FileSizeBytes { get; set; }

        [BsonElement("alerts")]
        public List<string> Alerts { get; set; } = new();

        [BsonElement("submittedAt")]
        public DateTime SubmittedAt { get; set; } = DateTime.UtcNow;
    }

    [BsonIgnoreExtraElements]
    public class SubmissionAlert
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("portalId")]
        public string PortalId { get; set; } = string.Empty;

        [BsonElement("alertType")]
        public string AlertType { get; set; } = string.Empty;

        [BsonElement("severity")]
        public string Severity { get; set; } = "Medium";

        [BsonElement("message")]
        public string Message { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("involvedStudentIds")]
        public List<string> InvolvedStudentIds { get; set; } = new();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("involvedSubmissionLogIds")]
        public List<string> InvolvedSubmissionLogIds { get; set; } = new();

        [BsonElement("metadata")]
        public BsonDocument? Metadata { get; set; }

        [BsonElement("isRead")]
        public bool IsRead { get; set; }

        [BsonElement("isDismissed")]
        public bool IsDismissed { get; set; }

        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}