using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class GradingConfigTestRun
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("gradingConfigId")]
        public string GradingConfigId { get; set; } = string.Empty;

        [BsonElement("gradingApiEndpoint")]
        public string GradingApiEndpoint { get; set; } = string.Empty;

        [BsonElement("configVersion")]
        public int ConfigVersion { get; set; }

        [BsonElement("usedOverride")]
        public bool UsedOverride { get; set; }

        [BsonElement("fileName")]
        public string FileName { get; set; } = string.Empty;

        [BsonElement("totalScore")]
        public decimal TotalScore { get; set; }

        [BsonElement("maxScore")]
        public decimal MaxScore { get; set; }

        [BsonElement("percentage")]
        public double Percentage { get; set; }

        [BsonElement("status")]
        public string Status { get; set; } = string.Empty;

        [BsonElement("createdAt")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("createdBy")]
        public string? CreatedBy { get; set; }

        [BsonElement("error")]
        public string? Error { get; set; }
    }
}