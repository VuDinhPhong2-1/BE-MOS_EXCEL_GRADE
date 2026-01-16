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

        [Range(0, 100)]
        [BsonElement("scoreValue")]
        public double? ScoreValue { get; set; }

        [StringLength(500)]
        [BsonElement("feedback")]
        public string? Feedback { get; set; }

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
}
