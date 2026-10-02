// MOS.ExcelGrading.Core/Models/BonusPoint.cs
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class BonusPoint
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string ClassId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("studentId")]
        public string StudentId { get; set; } = string.Empty;

        [BsonElement("date")]
        public DateTime Date { get; set; }

        /// <summary>Điểm cộng (có thể âm để trừ điểm)</summary>
        [BsonElement("points")]
        public double Points { get; set; }

        /// <summary>"participation" | "behavior" | "achievement" | custom</summary>
        [BsonElement("category")]
        public string Category { get; set; } = "participation";

        [BsonElement("reason")]
        public string? Reason { get; set; }

        /// <summary>Optional: gắn với buổi học cụ thể</summary>
        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("scheduleId")]
        public string? ScheduleId { get; set; }

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

    public static class BonusPointCategories
    {
        public const string Participation = "participation"; // Tham gia tích cực
        public const string Behavior = "behavior";           // Hành vi tốt
        public const string Achievement = "achievement";     // Đạt thành tích
    }
}
