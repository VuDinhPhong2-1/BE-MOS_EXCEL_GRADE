using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class GradingConfig
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonElement("gradingApiEndpoint")]
        public string GradingApiEndpoint { get; set; } = string.Empty;

        [BsonElement("subject")]
        public string Subject { get; set; } = string.Empty;

        [BsonElement("projectNumber")]
        public int ProjectNumber { get; set; }

        [BsonElement("projectCode")]
        public string ProjectCode { get; set; } = string.Empty;

        [BsonElement("displayName")]
        public string DisplayName { get; set; } = string.Empty;

        [BsonElement("status")]
        public string Status { get; set; } = GradingConfigStatuses.Draft;

        [BsonElement("version")]
        public int Version { get; set; } = 1;

        [BsonElement("normalizedMaxScore")]
        public decimal NormalizedMaxScore { get; set; } = GradingConfigDefaults.NormalizedMaxScore;

        [BsonElement("tasks")]
        public List<GradingConfigTask> Tasks { get; set; } = new();

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

        [BsonElement("publishedAt")]
        public DateTime? PublishedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("publishedBy")]
        public string? PublishedBy { get; set; }

        [BsonElement("summary")]
        public string? Summary { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class GradingConfigTask
    {
        [BsonElement("taskId")]
        public string TaskId { get; set; } = string.Empty;

        [BsonElement("taskName")]
        public string TaskName { get; set; } = string.Empty;

        [BsonElement("maxScore")]
        public decimal MaxScore { get; set; }

        [BsonElement("enabled")]
        public bool Enabled { get; set; } = true;

        [BsonElement("sortOrder")]
        public int SortOrder { get; set; }

        [BsonElement("ruleType")]
        public string RuleType { get; set; } = GradingRuleTypes.LegacyCode;

        [BsonElement("rule")]
        public BsonDocument Rule { get; set; } = new();
    }

    public static class GradingConfigStatuses
    {
        public const string Draft = "draft";
        public const string Active = "active";

        public static bool IsValid(string? value)
        {
            return string.Equals(value, Draft, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, Active, StringComparison.OrdinalIgnoreCase);
        }
    }

    public static class GradingRuleTypes
    {
        public const string LegacyCode = "legacy.code";
    }

    public static class GradingConfigDefaults
    {
        public const decimal NormalizedMaxScore = 125m;
    }
}