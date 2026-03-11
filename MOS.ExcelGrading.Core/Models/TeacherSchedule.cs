using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class TeacherSchedule
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [Required]
        public string OwnerId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        public string? SchoolId { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string? ClassId { get; set; }

        [StringLength(150)]
        public string ClassName { get; set; } = string.Empty;

        [Required]
        [StringLength(120)]
        public string Subject { get; set; } = string.Empty;

        [StringLength(100)]
        public string? RoomName { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string? RoomId { get; set; }

        [StringLength(60)]
        public string? PeriodLabel { get; set; }

        public DateTime Date { get; set; }

        [Required]
        [RegularExpression("^([01]\\d|2[0-3]):[0-5]\\d$", ErrorMessage = "Giờ bắt đầu phải theo định dạng HH:mm")]
        public string StartTime { get; set; } = string.Empty;

        [Required]
        [RegularExpression("^([01]\\d|2[0-3]):[0-5]\\d$", ErrorMessage = "Giờ kết thúc phải theo định dạng HH:mm")]
        public string EndTime { get; set; } = string.Empty;

        [StringLength(500)]
        public string? Notes { get; set; }

        public ScheduleReportBundle Reports { get; set; } = new();

        public bool IsActive { get; set; } = true;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? UpdatedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string? UpdatedBy { get; set; }
    }
}
