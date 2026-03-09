// MOS.ExcelGrading.Core/Models/Student.cs
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class Student
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [Required(ErrorMessage = "Họ và tên đệm là bắt buộc")]
        [StringLength(100)]
        public string MiddleName { get; set; } = string.Empty;

        [Required(ErrorMessage = "Tên là bắt buộc")]
        [StringLength(100)]
        public string FirstName { get; set; } = string.Empty;

        [Required]
        public string Status { get; set; } = "Active"; // Active, Inactive, Special, Auto Active

        [StringLength(1)]
        public string? CompetencyLevel { get; set; } // A, B, C, D

        [StringLength(500)]
        public string? Notes { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string? TeacherId { get; set; } // Lấy ID của user tạo student này

        [BsonRepresentation(BsonType.ObjectId)]
        public string? ClassId { get; set; }

        // Metadata
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        public string? CreatedBy { get; set; }

        public DateTime? UpdatedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string? UpdatedBy { get; set; }

        public bool IsActive { get; set; } = true;
    }
}
