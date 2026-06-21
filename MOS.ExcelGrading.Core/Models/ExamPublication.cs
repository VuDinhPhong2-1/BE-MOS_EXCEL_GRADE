using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.Models
{
    [BsonIgnoreExtraElements]
    public class ExamPublication
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = ObjectId.GenerateNewId().ToString();

        [Required]
        [StringLength(200)]
        [BsonElement("name")]
        public string Name { get; set; } = string.Empty;

        [StringLength(1000)]
        [BsonElement("description")]
        public string? Description { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("classId")]
        public string? ClassId { get; set; }

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("studentIds")]
        public List<string> StudentIds { get; set; } = new();

        [BsonElement("mode")]
        public string? Mode { get; set; }

        [BsonElement("startsAt")]
        public DateTime? StartsAt { get; set; }

        [BsonElement("endsAt")]
        public DateTime? EndsAt { get; set; }

        [BsonElement("durationMinutes")]
        public int? DurationMinutes { get; set; }

        [BsonElement("publicationToken")]
        public string PublicationToken { get; set; } = Guid.NewGuid().ToString("N");

        [BsonElement("projectSequence")]
        public List<ExamPublicationProject> ProjectSequence { get; set; } = new();

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("assignmentIds")]
        public List<string> AssignmentIds { get; set; } = new();

        [BsonElement("isActive")]
        public bool IsActive { get; set; } = true;

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

    public class ExamPublicationProject
    {
        [BsonElement("order")]
        public int Order { get; set; }

        [BsonElement("projectCode")]
        public string ProjectCode { get; set; } = string.Empty;

        [BsonRepresentation(BsonType.ObjectId)]
        [BsonElement("sourceAssignmentId")]
        public string? SourceAssignmentId { get; set; }

        [BsonElement("subject")]
        public string Subject { get; set; } = AssignmentFileSubjects.Excel;

        [BsonElement("templateFileName")]
        public string? TemplateFileName { get; set; }

        [BsonElement("instructionsFileName")]
        public string? InstructionsFileName { get; set; }

        [BsonElement("instructionsText")]
        public string? InstructionsText { get; set; }

        [BsonElement("helpFileName")]
        public string? HelpFileName { get; set; }

        [BsonElement("helpText")]
        public string? HelpText { get; set; }

        [BsonElement("gradingApiEndpoint")]
        public string GradingApiEndpoint { get; set; } = string.Empty;

        [BsonElement("taskSnapshot")]
        public List<ExamPublicationTaskSnapshotItem> TaskSnapshot { get; set; } = new();

        [BsonElement("modeRules")]
        public ExamPublicationModeRules? ModeRules { get; set; }
    }

    public class ExamPublicationTaskSnapshotItem
    {
        [BsonElement("taskId")]
        public string? TaskId { get; set; }

        [BsonElement("taskName")]
        public string? TaskName { get; set; }

        [BsonElement("maxScore")]
        public double? MaxScore { get; set; }

        [BsonElement("instructions")]
        public string? Instructions { get; set; }
    }

    public class ExamPublicationModeRules
    {
        [BsonElement("mode")]
        public string? Mode { get; set; }

        [BsonElement("showFeedback")]
        public bool? ShowFeedback { get; set; }

        [BsonElement("allowRestart")]
        public bool? AllowRestart { get; set; }

        [BsonElement("allowNextProject")]
        public bool? AllowNextProject { get; set; }

        [BsonElement("allowHelp")]
        public bool? AllowHelp { get; set; }
    }
}
