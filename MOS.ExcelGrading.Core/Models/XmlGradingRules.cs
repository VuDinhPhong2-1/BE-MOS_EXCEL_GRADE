using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
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
        public const string XmlMinOccurrences = "xmlMinOccurrences";

        public static readonly HashSet<string> Supported =
            new(StringComparer.OrdinalIgnoreCase)
            {
            XmlContainsNormalized,
            XmlContains,
            XmlEquivalentWholeFile,
            ExactStringContains,
            XmlMinOccurrences
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

    public static class SpecialConditionTypes
{
    public const string PictureBullet = "pictureBullet";
    public const string InsertedImage = "insertedImage"; // MỚI
    public const string ConvertTableToText = "convertTableToText";
    public const string Hyperlink = "hyperlink";

    public static readonly HashSet<string> Supported =
        new(StringComparer.OrdinalIgnoreCase)
        {
            PictureBullet,
            InsertedImage,
            ConvertTableToText,
            Hyperlink
        };
}

/// <summary>
/// Các chế độ ngắt dòng văn bản (Text Wrapping) của ảnh trong Word,
/// dùng để kiểm tra specialCondition = insertedImage.
/// </summary>
public static class ImageWrapTypes
{
    public const string Inline = "inline";
    public const string Square = "square";
    public const string Tight = "tight";
    public const string Through = "through";
    public const string TopAndBottom = "topAndBottom";
    public const string Behind = "behind";
    public const string InFront = "inFront";

    public static readonly HashSet<string> Supported = new(StringComparer.OrdinalIgnoreCase)
    {
        Inline, Square, Tight, Through, TopAndBottom, Behind, InFront
    };
}

    // [BsonIgnoreExtraElements] được thêm vào TẤT CẢ các class map với MongoDB
    // bên dưới để driver bỏ qua field lạ/thừa trong document thay vì throw
    // exception khi deserialize (đây là nguyên nhân gốc của lỗi
    // "Element 'specialCondition' does not match any field or property of
    // class XmlGradingCondition" — dữ liệu cũ trong DB có field specialCondition
    // bị lồng sai vị trí bên trong 1 condition thay vì ở cấp Task).

    [BsonIgnoreExtraElements]
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

    [BsonIgnoreExtraElements]
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

    [BsonIgnoreExtraElements]
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

        /// <summary>
        /// Điều kiện đặc biệt của riêng Task này — khớp đúng vị trí
        /// "task.specialCondition" phía FE (không nằm trong Condition).
        /// </summary>
        [BsonElement("specialCondition")]
        [JsonPropertyName("specialCondition")]
        public SpecialCondition? SpecialCondition { get; set; }
    }

    /// <summary>
    /// Khớp đúng interface SpecialCondition phía FE:
    /// { type: SpecialConditionType; config?: PictureBulletConfig }
    /// </summary>
    [BsonIgnoreExtraElements]
    public class SpecialCondition
    {
        [BsonElement("type")]
        [JsonPropertyName("type")]
        public string Type { get; set; } = string.Empty;

        [BsonElement("score")]
        [JsonPropertyName("score")]
        public decimal Score { get; set; }

        [BsonElement("feedback")]
        [JsonPropertyName("feedback")]
        public ConditionFeedback Feedback { get; set; } = new();

        [BsonElement("config")]
        [JsonPropertyName("config")]
        public PictureBulletConfig? Config { get; set; }

        // MỚI: dùng riêng cho type = insertedImage
        [BsonElement("imageInsertConfig")]
        [JsonPropertyName("imageInsertConfig")]
        public ImageInsertConfig? ImageInsertConfig { get; set; }

        [BsonElement("convertTableToTextConfig")]
        [JsonPropertyName("convertTableToTextConfig")]
        public ConvertTableToTextConfig? ConvertTableToTextConfig { get; set; }

        [BsonElement("hyperlinkConfig")]
        [JsonPropertyName("hyperlinkConfig")]
        public HyperlinkConfig? HyperlinkConfig { get; set; }
}

    [BsonIgnoreExtraElements]
    public class HyperlinkConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("relsFile")]
        [JsonPropertyName("relsFile")]
        public string? RelsFile { get; set; } = "word/_rels/document.xml.rels";

        [BsonElement("displayText")]
        [JsonPropertyName("displayText")]
        public string? DisplayText { get; set; }

        [BsonElement("url")]
        [JsonPropertyName("url")]
        public string? Url { get; set; }

        [BsonElement("caseSensitiveText")]
        [JsonPropertyName("caseSensitiveText")]
        public bool? CaseSensitiveText { get; set; } = false;
    }

    [BsonIgnoreExtraElements]
    public class ConvertTableToTextConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("anchorText")]
        [JsonPropertyName("anchorText")]
        public string? AnchorText { get; set; }

        [BsonElement("expectedRows")]
        [JsonPropertyName("expectedRows")]
        public List<string> ExpectedRows { get; set; } = new();

        [BsonElement("minRows")]
        [JsonPropertyName("minRows")]
        public int? MinRows { get; set; }

        [BsonElement("minTabsPerRow")]
        [JsonPropertyName("minTabsPerRow")]
        public int? MinTabsPerRow { get; set; }

        [BsonElement("requireNoTables")]
        [JsonPropertyName("requireNoTables")]
        public bool? RequireNoTables { get; set; } = true;
    }

    /// <summary>
    /// Khớp đúng interface PictureBulletConfig phía FE:
    /// { level?: number; assetId?: string; imageHash?: string }
    /// assetId chỉ mang tính tham chiếu (id ảnh đã upload), không dùng để chấm.
    /// imageHash là SHA256 dùng để so khớp khi chấm.
    /// </summary>
    [BsonIgnoreExtraElements]
    public class PictureBulletConfig
    {
        [BsonElement("level")]
        [JsonPropertyName("level")]
        public int? Level { get; set; }

        [BsonElement("assetId")]
        [JsonPropertyName("assetId")]
        public string? AssetId { get; set; }

        [BsonElement("imageHash")]
        [JsonPropertyName("imageHash")]
        public string? ImageHash { get; set; }

        // MỚI: dHash 64-bit dạng hex 16 ký tự, dùng để so sánh ảnh chịu được
        // nén JPEG lại khi Word lưu file. Optional để tương thích ruleset cũ
        // (chưa có field này thì fallback về so ImageHash tuyệt đối).
        [BsonElement("perceptualHash")]
        [JsonPropertyName("perceptualHash")]
        public string? PerceptualHash { get; set; }
    }

    /// <summary>
    /// Khớp đúng interface ImageInsertConfig phía FE:
    /// { assetId?: string; imageHash?: string; wrapType?: ImageWrapType }
    /// assetId chỉ mang tính tham chiếu (id ảnh đã upload), không dùng để chấm.
    /// imageHash là SHA256 dùng để so khớp khi chấm.
    /// wrapType để trống nếu không cần kiểm tra chế độ ngắt dòng.
    /// </summary>
    [BsonIgnoreExtraElements]
    public class ImageInsertConfig
    {
        [BsonElement("assetId")]
        [JsonPropertyName("assetId")]
        public string? AssetId { get; set; }

        [BsonElement("imageHash")]
        [JsonPropertyName("imageHash")]
        public string? ImageHash { get; set; }

        [BsonElement("wrapType")]
        [JsonPropertyName("wrapType")]
        public string? WrapType { get; set; }

        // MỚI: dHash 64-bit dạng hex 16 ký tự, dùng để so sánh ảnh chịu được
        // nén JPEG lại khi Word lưu file. Optional để tương thích ruleset cũ
        // (chưa có field này thì fallback về so ImageHash tuyệt đối).
        [BsonElement("perceptualHash")]
        [JsonPropertyName("perceptualHash")]
        public string? PerceptualHash { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class XmlExpectedVariant
    {
        [JsonPropertyName("expectedValues")]
        [BsonElement("expectedValues")]
        public List<string> ExpectedValues { get; set; } = new();
    }

    [BsonIgnoreExtraElements]
    public class XmlGradingCondition
    {
        [BsonElement("conditionId")]
        public string ConditionId { get; set; } = string.Empty;

        [BsonElement("score")]
        public decimal Score { get; set; }

        [BsonElement("sourceFile")]
        public string SourceFile { get; set; } = string.Empty;

        [JsonPropertyName("expectedVariants")]
        [BsonElement("expectedVariants")]
        public List<XmlExpectedVariant> ExpectedVariants { get; set; } = new();

        [JsonPropertyName("ignoreAttributes")]
        [BsonElement("ignoreAttributes")]
        public List<string> IgnoreAttributes { get; set; } = new();

        [BsonElement("compareMode")]
        public string CompareMode { get; set; } = XmlGradingCompareModes.XmlContainsNormalized;

        [BsonElement("matchPolicy")]
        public string MatchPolicy { get; set; } = XmlGradingMatchPolicies.All;

        [JsonPropertyName("minOccurrences")]
        [BsonElement("minOccurrences")]
        public int? MinOccurrences { get; set; }

        [JsonPropertyName("maxOccurrences")]
        [BsonElement("maxOccurrences")]
        public int? MaxOccurrences { get; set; }

        [BsonElement("feedback")]
        public ConditionFeedback Feedback { get; set; } = new();

        [BsonElement("stopTaskIfFailed")]
        public bool StopTaskIfFailed { get; set; } = false;
    }

    [BsonIgnoreExtraElements]
    public class ConditionFeedback
    {
        [BsonElement("successDetail")]
        public string SuccessDetail { get; set; } = string.Empty;

        [BsonElement("errorMessage")]
        public string ErrorMessage { get; set; } = string.Empty;

        [BsonElement("fixAction")]
        public string FixAction { get; set; } = string.Empty;
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
    }

    public class XmlRuleValidationResult
    {
        public bool IsValid => Errors.Count == 0;
        public List<string> Errors { get; set; } = new();
        public List<string> Warnings { get; set; } = new();
    }
}
