using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    public class GradingAttempt
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonElement("projectEndpoint")]
        public string ProjectEndpoint { get; set; } = string.Empty;

        [BsonElement("projectId")]
        public string ProjectId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string? ClassId { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("assignmentId")]
        public string? AssignmentId { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("studentId")]
        public string? StudentId { get; set; }

        [BsonElement("totalScore")]
        public decimal TotalScore { get; set; }

        [BsonElement("maxScore")]
        public decimal MaxScore { get; set; }

        [BsonElement("percentage")]
        public double Percentage { get; set; }

        [BsonElement("status")]
        public string Status { get; set; } = string.Empty;

        [BsonElement("taskResults")]
        public List<GradingAttemptTask> TaskResults { get; set; } = new();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("gradedBy")]
        public string? GradedBy { get; set; }

        [BsonElement("gradedAt")]
        public DateTime GradedAt { get; set; } = DateTime.UtcNow;
    }

    public class GradingAttemptTask
    {
        [BsonElement("taskId")]
        public string TaskId { get; set; } = string.Empty;

        [BsonElement("taskName")]
        public string TaskName { get; set; } = string.Empty;

        [BsonElement("score")]
        public decimal Score { get; set; }

        [BsonElement("maxScore")]
        public decimal MaxScore { get; set; }

        [BsonElement("isPassed")]
        public bool IsPassed { get; set; }

        [BsonElement("errorCount")]
        public int ErrorCount { get; set; }
    }
}
