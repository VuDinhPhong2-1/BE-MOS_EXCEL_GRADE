// MOS.ExcelGrading.Core/Models/Score.cs
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class Score
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [Required]
        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("studentId")]
        public string StudentId { get; set; } = string.Empty;

        [Required]
        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("assignmentId")]
        public string AssignmentId { get; set; } = string.Empty;

        [Required]
        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string ClassId { get; set; } = string.Empty;

        [Range(0, 1000)]
        [BsonElement("scoreValue")]
        public double? ScoreValue { get; set; }

        [StringLength(500)]
        [BsonElement("feedback")]
        public string? Feedback { get; set; }

        [BsonElement("autoGradingErrors")]
        public List<string> AutoGradingErrors { get; set; } = new();

        [BsonElement("autoGradingTaskResults")]
        public List<ScoreTaskResult> AutoGradingTaskResults { get; set; } = new();

        [BsonElement("gradedAt")]
        public DateTime? GradedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("gradedBy")]
        public string? GradedBy { get; set; }

        // ========== METADATA ==========
        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("createdBy")]
        public string? CreatedBy { get; set; }

        [BsonElement("updatedAt")]
        public DateTime? UpdatedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("updatedBy")]
        public string? UpdatedBy { get; set; }
    }

    public class ScoreTaskResult
    {
        [BsonElement("taskId")]
        public string TaskId { get; set; } = string.Empty;

        [BsonElement("taskName")]
        public string TaskName { get; set; } = string.Empty;

        [BsonElement("score")]
        public double Score { get; set; }

        [BsonElement("maxScore")]
        public double MaxScore { get; set; }

        [BsonElement("isPassed")]
        public bool IsPassed { get; set; }

        [BsonElement("details")]
        public List<string> Details { get; set; } = new();

        [BsonElement("errors")]
        public List<string> Errors { get; set; } = new();

        [BsonElement("fixActions")]
        public List<string> FixActions { get; set; } = new();

        [BsonElement("displayIssues")]
        public List<ScoreTaskDisplayIssue> DisplayIssues { get; set; } = new();
    }

    public class ScoreTaskDisplayIssue
    {
        [BsonElement("heading")]
        public string Heading { get; set; } = string.Empty;

        [BsonElement("message")]
        public string Message { get; set; } = string.Empty;

        [BsonElement("fixAction")]
        public string FixAction { get; set; } = string.Empty;
    }
}
