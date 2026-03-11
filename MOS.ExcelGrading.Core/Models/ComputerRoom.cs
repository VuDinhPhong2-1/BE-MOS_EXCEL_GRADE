using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class ComputerRoom
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string? Id { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [Required]
        public string SchoolId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [Required]
        public string OwnerId { get; set; } = string.Empty;

        [Required]
        [StringLength(100, MinimumLength = 2)]
        public string Name { get; set; } = string.Empty;

        [Range(0, 1000)]
        public int StudentMachineCount { get; set; }

        [Range(0, 50)]
        public int TeacherMachineCount { get; set; } = 1;

        [Range(0, 1000)]
        public int BrokenMachineCount { get; set; } = 0;

        [StringLength(120)]
        public string NetSupportStatus { get; set; } = "Tốt";

        [StringLength(120)]
        public string AudioStatus { get; set; } = "Tốt";

        [StringLength(120)]
        public string CoolingStatus { get; set; } = "Tốt";

        [StringLength(120)]
        public string DevicesPoweredOffStatus { get; set; } = "Rồi";

        [StringLength(120)]
        public string SeatingOrderStatus { get; set; } = "Tốt";

        [StringLength(120)]
        public string RoomHygieneStatus { get; set; } = "Tốt";

        public bool IsActive { get; set; } = true;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? UpdatedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        public string? UpdatedBy { get; set; }
    }
}
