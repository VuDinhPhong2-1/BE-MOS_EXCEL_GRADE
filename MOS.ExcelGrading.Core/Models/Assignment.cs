// MOS.ExcelGrading.Core/Models/Assignment.cs
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    public class Assignment
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [Required(ErrorMessage = "Assignment name is required")]
        [StringLength(200)]
        [BsonElement("name")]
        public string Name { get; set; } = string.Empty;

        [StringLength(1000)]
        [BsonElement("description")]
        public string? Description { get; set; }

        [Required]
        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string ClassId { get; set; } = string.Empty;

        [Required]
        [Range(0, 100)]
        [BsonElement("maxScore")]
        public double MaxScore { get; set; } = 10;


        // ========== THÊM MỚI: LIÊN KẾT VỚI GRADING API ==========
        /// <summary>
        /// API endpoint để chấm điểm (vd: "project09", "project10", "project11")
        /// </summary>
        [BsonElement("gradingApiEndpoint")]
        public string? GradingApiEndpoint { get; set; }

        /// <summary>
        /// Loại bài tập: "auto" (tự động chấm), "manual" (chấm thủ công)
        /// </summary>
        [BsonElement("gradingType")]
        public string GradingType { get; set; } = "manual"; // "auto" | "manual"

        [BsonElement("isActive")]
        public bool IsActive { get; set; } = true;

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

    // ========== ĐỊNH NGHĨA GRADING TYPES ==========
    public static class GradingTypes
    {
        public const string Auto = "auto";
        public const string Manual = "manual";
    }

    // ========== ĐỊNH NGHĨA GRADING API ENDPOINTS ==========
    public static class GradingApiEndpoints
    {
        public const string Project01 = "project01";
        public const string Project02 = "project02";
        public const string Project03 = "project03";
        public const string Project04 = "project04";
        public const string Project05 = "project05";
        public const string Project06 = "project06";
        public const string Project07 = "project07";
        public const string Project08 = "project08";
        public const string Project09 = "project09";
        public const string Project10 = "project10";
        public const string Project11 = "project11";
        public const string Project12 = "project12";
        public const string Project13 = "project13";
        public const string Project14 = "project14";
        public const string Project15 = "project15";
        public const string Project16 = "project16";

        public static List<string> GetAllEndpoints() => new()
        {
            Project01,
            Project02,
            Project03,
            Project04,
            Project05,
            Project06,
            Project07,
            Project08,
            Project09,
            Project10,
            Project11,
            Project12,
            Project13,
            Project14,
            Project15,
            Project16
        };


        public static bool IsValidEndpoint(string endpoint) =>
            GetAllEndpoints().Contains(endpoint);
    }
}
