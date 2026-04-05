using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class GradingTestBugNote
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("userId")]
        public string UserId { get; set; } = string.Empty;

        [BsonElement("projectCode")]
        public string ProjectCode { get; set; } = string.Empty;

        [BsonElement("projectDisplayName")]
        public string ProjectDisplayName { get; set; } = string.Empty;

        [BsonElement("title")]
        public string Title { get; set; } = string.Empty;

        [BsonElement("description")]
        public string Description { get; set; } = string.Empty;

        [BsonElement("severity")]
        public string Severity { get; set; } = "medium";

        [BsonElement("scoreSummary")]
        public GradingTestBugNoteScoreSummary? ScoreSummary { get; set; }

        [BsonElement("gradingError")]
        public string? GradingError { get; set; }

        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonElement("updatedAt")]
        public DateTime? UpdatedAt { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class GradingTestBugNoteScoreSummary
    {
        [BsonElement("totalScore")]
        public double TotalScore { get; set; }

        [BsonElement("maxScore")]
        public double MaxScore { get; set; }

        [BsonElement("percentage")]
        public double Percentage { get; set; }

        [BsonElement("status")]
        public string Status { get; set; } = string.Empty;
    }
}
