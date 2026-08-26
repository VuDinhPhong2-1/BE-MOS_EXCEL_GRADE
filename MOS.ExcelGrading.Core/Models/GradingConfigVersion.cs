using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class GradingConfigVersion
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("gradingConfigId")]
        public string GradingConfigId { get; set; } = string.Empty;

        [BsonElement("version")]
        public int Version { get; set; }

        [BsonElement("action")]
        public string Action { get; set; } = GradingConfigVersionActions.Publish;

        [BsonElement("summary")]
        public string? Summary { get; set; }

        [BsonElement("snapshot")]
        public GradingConfig Snapshot { get; set; } = new();

        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("createdBy")]
        public string? CreatedBy { get; set; }
    }

    public static class GradingConfigVersionActions
    {
        public const string ImportFromCode = "import-from-code";
        public const string Publish = "publish";
        public const string Restore = "restore";
    }
}