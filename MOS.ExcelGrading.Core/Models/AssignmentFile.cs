using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace MOS.ExcelGrading.Core.Models
{
    public class AssignmentFile
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("assignmentId")]
        public string AssignmentId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string ClassId { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("gridFsFileId")]
        public string GridFsFileId { get; set; } = string.Empty;

        [BsonElement("originalName")]
        public string OriginalName { get; set; } = string.Empty;

        [BsonElement("extension")]
        public string Extension { get; set; } = string.Empty;

        [BsonElement("contentType")]
        public string ContentType { get; set; } = "application/octet-stream";

        [BsonElement("sizeBytes")]
        public long SizeBytes { get; set; }

        [BsonElement("sha256")]
        public string Sha256 { get; set; } = string.Empty;

        [BsonElement("subject")]
        public string Subject { get; set; } = AssignmentFileSubjects.Excel;

        [BsonElement("kind")]
        public string Kind { get; set; } = AssignmentFileKinds.Template;

        [BsonElement("version")]
        public int Version { get; set; } = 1;

        [BsonElement("isActive")]
        public bool IsActive { get; set; } = true;

        [BsonElement("uploadedAt")]
        public DateTime UploadedAt { get; set; } = DateTime.UtcNow;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("uploadedBy")]
        public string? UploadedBy { get; set; }

        [BsonElement("updatedAt")]
        public DateTime? UpdatedAt { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("updatedBy")]
        public string? UpdatedBy { get; set; }
    }

    public static class AssignmentFileSubjects
    {
        public const string Excel = "excel";
        public const string Word = "word";
        public const string Ppt = "ppt";

        public static string Normalize(string? value)
        {
            var normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
            if (normalized == "powerpoint")
            {
                return Ppt;
            }

            return normalized;
        }

        public static bool IsValid(string? value)
        {
            var normalized = Normalize(value);
            return normalized == Excel || normalized == Word || normalized == Ppt;
        }
    }

    public static class AssignmentFileKinds
    {
        public const string Template = "template";
        public const string Answer = "answer";
        public const string Instructions = "instructions";
        public const string Help = "help";

        public static string Normalize(string? value) =>
            (value ?? string.Empty).Trim().ToLowerInvariant();

        public static bool IsValid(string? value)
        {
            var normalized = Normalize(value);
            return normalized == Template ||
                normalized == Answer ||
                normalized == Instructions ||
                normalized == Help;
        }
    }
}
