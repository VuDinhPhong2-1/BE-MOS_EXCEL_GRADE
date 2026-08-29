using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MOS.ExcelGrading.Core.Models
{
    public static class XmlGradingCompareModes
    {
        // Bỏ qua khác biệt format/whitespace của XML
        public const string XmlContainsNormalized = "xmlContainsNormalized";

        // So chuỗi XML nguyên văn, chỉ Trim đầu/cuối expected
        public const string XmlContains = "xmlContains";

        // So sánh tương đương toàn bộ XML
        public const string XmlEquivalentWholeFile = "xmlEquivalentWholeFile";

        // Tìm chuỗi tuyệt đối trong raw XML
        public const string ExactStringContains = "exactStringContains";

        public static readonly HashSet<string> Supported =
            new(StringComparer.OrdinalIgnoreCase)
            {
            XmlContainsNormalized,
            XmlContains,
            XmlEquivalentWholeFile,
            ExactStringContains
            };
    }

    public static class XmlGradingMatchPolicies
    {
        public const string All = "all";
        public const string Any = "any";
        public const string Ordered = "ordered";

        public static readonly HashSet<string> Supported = new(StringComparer.OrdinalIgnoreCase)
        {
            All,
            Any,
            Ordered
        };
    }

    public class GradingRuleSet
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = string.Empty;

        [BsonElement("subject")]
        public string Subject { get; set; } = string.Empty;

        [BsonElement("version")]
        public string Version { get; set; } = string.Empty;

        [BsonElement("isActive")]
        public bool IsActive { get; set; } = true;

        [BsonElement("projects")]
        public List<ProjectXmlRule> Projects { get; set; } = new();
    }

    public class ProjectXmlRule
    {
        [BsonElement("projectCode")]
        public string ProjectCode { get; set; } = string.Empty;

        [BsonElement("projectName")]
        public string ProjectName { get; set; } = string.Empty;

        [BsonElement("maxScore")]
        public decimal MaxScore { get; set; } = 125m;

        [BsonElement("tasks")]
        public List<TaskXmlRule> Tasks { get; set; } = new();
    }

    public class TaskXmlRule
    {
        [BsonElement("taskId")]
        public string TaskId { get; set; } = string.Empty;

        [BsonElement("taskName")]
        public string TaskName { get; set; } = string.Empty;

        [BsonElement("maxScore")]
        public decimal MaxScore { get; set; }

        [BsonElement("conditions")]
        public List<XmlGradingCondition> Conditions { get; set; } = new();
    }

    public class XmlGradingCondition
    {
        [BsonElement("conditionId")]
        public string ConditionId { get; set; } = string.Empty;

        [BsonElement("score")]
        public decimal Score { get; set; }

        [BsonElement("specialCondition")]
        [JsonPropertyName("specialCondition")]
        public SpecialCondition? SpecialCondition { get; set; }

        [BsonElement("sourceFile")]
        public string SourceFile { get; set; } = string.Empty;

        [JsonPropertyName("expectedValues")]
        [BsonElement("expectedValues")]
        public List<string> ExpectedValues { get; set; } = new();

        [BsonElement("compareMode")]
        public string CompareMode { get; set; } = XmlGradingCompareModes.XmlContainsNormalized;

        [BsonElement("matchPolicy")]
        public string MatchPolicy { get; set; } = XmlGradingMatchPolicies.All;

        [BsonElement("feedback")]
        public ConditionFeedback Feedback { get; set; } = new();

        [BsonElement("stopTaskIfFailed")]
        public bool StopTaskIfFailed { get; set; } = false;
    }
    public class ConditionFeedback
    {
        [BsonElement("successDetail")]
        public string SuccessDetail { get; set; } = string.Empty;

        [BsonElement("errorMessage")]
        public string ErrorMessage { get; set; } = string.Empty;

        [BsonElement("fixAction")]
        public string FixAction { get; set; } = string.Empty;
    }

    public enum SpecialConditionType
{
    None = 0,
    PictureBullet = 1
}

public class SpecialCondition
{
    [BsonElement("type")]
    [JsonPropertyName("type")]
    public SpecialConditionType Type { get; set; }

    [BsonElement("pictureBullet")]
    [JsonPropertyName("pictureBullet")]
    public PictureBulletConfig? PictureBullet { get; set; }
}

public class PictureBulletConfig
{
    /// <summary>SHA256 của hình ảnh bullet chuẩn (bắt buộc).</summary>
    [BsonElement("expectedImageSha256")]
    [JsonPropertyName("expectedImageSha256")]
    public string ExpectedImageSha256 { get; set; } = string.Empty;

    /// <summary>File chứa paragraph cần kiểm tra. Mặc định word/document.xml.</summary>
    [BsonElement("documentPart")]
    [JsonPropertyName("documentPart")]
    public string DocumentPart { get; set; } = "word/document.xml";

    /// <summary>Index paragraph (đếm toàn bộ w:p, kể cả không có numPr). null = kiểm tra mọi paragraph.</summary>
    [BsonElement("paragraphIndex")]
    [JsonPropertyName("paragraphIndex")]
    public int? ParagraphIndex { get; set; }

    /// <summary>Giới hạn level (ilvl). null = mọi level.</summary>
    [BsonElement("level")]
    [JsonPropertyName("level")]
    public int? Level { get; set; }

    /// <summary>Giới hạn numId. null = không giới hạn.</summary>
    [BsonElement("numId")]
    [JsonPropertyName("numId")]
    public int? NumId { get; set; }

    [BsonElement("requirePictureBullet")]
    [JsonPropertyName("requirePictureBullet")]
    public bool RequirePictureBullet { get; set; } = true;
}

public class SpecialConditionEvaluationResult
{
    public bool IsPassed { get; set; }
    public string Type { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string? ExpectedSha256 { get; set; }
    public string? ActualSha256 { get; set; }
    public string? ImagePath { get; set; }
    public int? NumPicBulletId { get; set; }
    public string? RelationshipId { get; set; }
}

    public class ExpectedValueJsonConverter : JsonConverter<List<string>>
    {
        public override List<string> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.String)
            {
                var value = reader.GetString();
                return string.IsNullOrWhiteSpace(value) ? new List<string>() : new List<string> { value };
            }

            if (reader.TokenType == JsonTokenType.StartArray)
            {
                var values = JsonSerializer.Deserialize<List<string>>(ref reader, options) ?? new List<string>();
                return values.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            }

            throw new JsonException("expectedValue must be a non-empty string or an array of non-empty strings.");
        }

        public override void Write(Utf8JsonWriter writer, List<string> value, JsonSerializerOptions options)
        {
            if (value.Count == 1)
            {
                writer.WriteStringValue(value[0]);
                return;
            }

            JsonSerializer.Serialize(writer, value, options);
        }
    }

    public class ExpectedMatchResult
    {
        public string ExpectedValue { get; set; } = string.Empty;
        public bool IsMatched { get; set; }
        public int? MatchIndex { get; set; }
    }

    public class XmlConditionEvaluationResult
    {
        public string ConditionId { get; set; } = string.Empty;
        public string SourceFile { get; set; } = string.Empty;
        public string CompareMode { get; set; } = XmlGradingCompareModes.XmlContainsNormalized;
        public string MatchPolicy { get; set; } = XmlGradingMatchPolicies.All;
        public decimal ScoreAwarded { get; set; }
        public decimal MaxConditionScore { get; set; }
        public bool IsPassed { get; set; }
        public List<string> MatchedExpectedValues { get; set; } = new();
        public List<string> MissingExpectedValues { get; set; } = new();
        public ConditionFeedback Feedback { get; set; } = new();
        public SpecialConditionEvaluationResult? SpecialConditionResult { get; set; }
    }

    public class XmlRuleValidationResult
    {
        public bool IsValid => Errors.Count == 0;
        public List<string> Errors { get; set; } = new();
        public List<string> Warnings { get; set; } = new();
    }
}