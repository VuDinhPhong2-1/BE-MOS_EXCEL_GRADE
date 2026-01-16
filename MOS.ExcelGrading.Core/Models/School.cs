using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class School
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [Required(ErrorMessage = "Tên trường là bắt buộc")]
        [StringLength(200, MinimumLength = 3)]
        public string Name { get; set; } = string.Empty;

        [Required(ErrorMessage = "Mã trường là bắt buộc")]
        [StringLength(50)]
        public string Code { get; set; } = string.Empty;

        [StringLength(500)]
        public string? Address { get; set; }

        [Phone]
        public string? PhoneNumber { get; set; }

        [EmailAddress]
        public string? Email { get; set; }

        public string? Website { get; set; }
        public string? Description { get; set; }
        public string? Logo { get; set; }

        // ========== PHÂN QUYỀN ==========
        /// <summary>
        /// ID của Teacher tạo ra school này (Owner)
        /// Chỉ owner mới có quyền CRUD school này
        /// </summary>
        [BsonRepresentation(BsonType.ObjectId)]
        [Required]
        public string OwnerId { get; set; } = string.Empty;

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
