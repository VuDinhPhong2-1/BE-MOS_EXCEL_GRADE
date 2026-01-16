using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class Class
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [Required(ErrorMessage = "Tên lớp là bắt buộc")]
        [StringLength(100, MinimumLength = 2)]
        public string Name { get; set; } = string.Empty;

        // ========== QUAN HỆ VỚI SCHOOL ==========
        [Required(ErrorMessage = "SchoolId là bắt buộc")]
        [BsonRepresentation(BsonType.ObjectId)]
        public string SchoolId { get; set; } = string.Empty;

        // ========== PHÂN QUYỀN ==========
        /// <summary>
        /// ID của Teacher tạo ra class này (Owner)
        /// Chỉ owner mới có quyền CRUD class này
        /// </summary>
        [BsonRepresentation(BsonType.ObjectId)]
        [Required]
        public string OwnerId { get; set; } = string.Empty;

        // ========== THÔNG TIN LỚP ==========
        public string? Description { get; set; }
        public int? MaxStudents { get; set; }
        public int CurrentStudents { get; set; } = 0;
        public string? AcademicYear { get; set; }
        public string? Grade { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public List<string> StudentIds { get; set; } = new List<string>();

        // ========== METADATA ==========
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        public string? CreatedBy { get; set; }

        public DateTime? UpdatedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string? UpdatedBy { get; set; }

        public bool IsActive { get; set; } = true;
    }
}
