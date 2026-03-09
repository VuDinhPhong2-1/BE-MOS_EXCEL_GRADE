using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    public class StudentScheduleAttendance
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string OwnerId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        public string ScheduleId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        public string? SchoolId { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string ClassId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        public string StudentId { get; set; } = string.Empty;

        public DateTime Date { get; set; }

        public string Status { get; set; } = AttendanceStatus.Present;

        public string? Note { get; set; }

        public DateTime MarkedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        public string? MarkedBy { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public DateTime? UpdatedAt { get; set; }
    }

    public static class AttendanceStatus
    {
        public const string Present = "Present";
        public const string Absent = "Absent";
    }
}
