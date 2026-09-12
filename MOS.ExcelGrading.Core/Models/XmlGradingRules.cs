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
    public const string SectionBreakBeforeText = "sectionBreakBeforeText";
    public const string PictureStyle = "pictureStyle";
    public const string TextBoxContainsText = "textBoxContainsText";
    public const string PageMargins = "pageMargins";
    public const string DocumentStyleSet = "documentStyleSet";
    public const string PageBorder = "pageBorder";
    public const string ExcelTableName = "excelTableName";
    public const string ExcelWorksheetPageSetup = "excelWorksheetPageSetup";
    public const string ExcelClearCellFormatting = "excelClearCellFormatting";
    public const string ExcelDataModelImport = "excelDataModelImport";
    public const string ExcelCompatibilityReport = "excelCompatibilityReport";
    public const string ExcelMergedRange = "excelMergedRange";
    public const string ExcelCellHyperlink = "excelCellHyperlink";
    public const string ExcelIconSetConditionalFormatting = "excelIconSetConditionalFormatting";
    public const string ExcelChartDataRange = "excelChartDataRange";
    public const string ExcelChartStyle = "excelChartStyle";

    public static readonly HashSet<string> Supported =
        new(StringComparer.OrdinalIgnoreCase)
        {
            PictureBullet,
            InsertedImage,
            ConvertTableToText,
            Hyperlink,
            SectionBreakBeforeText,
            PictureStyle,
            TextBoxContainsText,
            PageMargins,
            DocumentStyleSet,
            PageBorder,
            ExcelTableName,
            ExcelWorksheetPageSetup,
            ExcelClearCellFormatting,
            ExcelDataModelImport,
            ExcelCompatibilityReport,
            ExcelMergedRange,
            ExcelCellHyperlink,
            ExcelIconSetConditionalFormatting,
            ExcelChartDataRange,
            ExcelChartStyle
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

    public class GradingRuleSetSummary
    {
        public string Id { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
        public bool IsActive { get; set; }
        public int ProjectCount { get; set; }
        public int TaskCount { get; set; }
        public int ConditionCount { get; set; }
        public decimal MaxScore { get; set; }
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

        [BsonElement("sectionBreakBeforeTextConfig")]
        [JsonPropertyName("sectionBreakBeforeTextConfig")]
        public SectionBreakBeforeTextConfig? SectionBreakBeforeTextConfig { get; set; }

        [BsonElement("pictureStyleConfig")]
        [JsonPropertyName("pictureStyleConfig")]
        public PictureStyleConfig? PictureStyleConfig { get; set; }

        [BsonElement("textBoxContainsTextConfig")]
        [JsonPropertyName("textBoxContainsTextConfig")]
        public TextBoxContainsTextConfig? TextBoxContainsTextConfig { get; set; }

        [BsonElement("pageMarginsConfig")]
        [JsonPropertyName("pageMarginsConfig")]
        public PageMarginsConfig? PageMarginsConfig { get; set; }

        [BsonElement("documentStyleSetConfig")]
        [JsonPropertyName("documentStyleSetConfig")]
        public DocumentStyleSetConfig? DocumentStyleSetConfig { get; set; }

        [BsonElement("pageBorderConfig")]
        [JsonPropertyName("pageBorderConfig")]
        public PageBorderConfig? PageBorderConfig { get; set; }

        [BsonElement("excelTableNameConfig")]
        [JsonPropertyName("excelTableNameConfig")]
        public ExcelTableNameConfig? ExcelTableNameConfig { get; set; }

        [BsonElement("excelWorksheetPageSetupConfig")]
        [JsonPropertyName("excelWorksheetPageSetupConfig")]
        public ExcelWorksheetPageSetupConfig? ExcelWorksheetPageSetupConfig { get; set; }

        [BsonElement("excelClearCellFormattingConfig")]
        [JsonPropertyName("excelClearCellFormattingConfig")]
        public ExcelClearCellFormattingConfig? ExcelClearCellFormattingConfig { get; set; }

        [BsonElement("excelDataModelImportConfig")]
        [JsonPropertyName("excelDataModelImportConfig")]
        public ExcelDataModelImportConfig? ExcelDataModelImportConfig { get; set; }

        [BsonElement("excelCompatibilityReportConfig")]
        [JsonPropertyName("excelCompatibilityReportConfig")]
        public ExcelCompatibilityReportConfig? ExcelCompatibilityReportConfig { get; set; }

        [BsonElement("excelMergedRangeConfig")]
        [JsonPropertyName("excelMergedRangeConfig")]
        public ExcelMergedRangeConfig? ExcelMergedRangeConfig { get; set; }

        [BsonElement("excelCellHyperlinkConfig")]
        [JsonPropertyName("excelCellHyperlinkConfig")]
        public ExcelCellHyperlinkConfig? ExcelCellHyperlinkConfig { get; set; }

        [BsonElement("excelIconSetConditionalFormattingConfig")]
        [JsonPropertyName("excelIconSetConditionalFormattingConfig")]
        public ExcelIconSetConditionalFormattingConfig? ExcelIconSetConditionalFormattingConfig { get; set; }

        [BsonElement("excelChartDataRangeConfig")]
        [JsonPropertyName("excelChartDataRangeConfig")]
        public ExcelChartDataRangeConfig? ExcelChartDataRangeConfig { get; set; }

        [BsonElement("excelChartStyleConfig")]
        [JsonPropertyName("excelChartStyleConfig")]
        public ExcelChartStyleConfig? ExcelChartStyleConfig { get; set; }
}

    [BsonIgnoreExtraElements]
    public class ExcelMergedRangeConfig
    {
        [BsonElement("worksheetName")]
        [JsonPropertyName("worksheetName")]
        public string? WorksheetName { get; set; }

        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; }

        [BsonElement("range")]
        [JsonPropertyName("range")]
        public string? Range { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class ExcelCellHyperlinkConfig
    {
        [BsonElement("worksheetName")]
        [JsonPropertyName("worksheetName")]
        public string? WorksheetName { get; set; }

        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; }

        [BsonElement("cell")]
        [JsonPropertyName("cell")]
        public string? Cell { get; set; }

        [BsonElement("location")]
        [JsonPropertyName("location")]
        public string? Location { get; set; }

        [BsonElement("target")]
        [JsonPropertyName("target")]
        public string? Target { get; set; }

        [BsonElement("display")]
        [JsonPropertyName("display")]
        public string? Display { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class ExcelIconSetConditionalFormattingConfig
    {
        [BsonElement("worksheetName")]
        [JsonPropertyName("worksheetName")]
        public string? WorksheetName { get; set; }

        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; }

        [BsonElement("range")]
        [JsonPropertyName("range")]
        public string? Range { get; set; }

        [BsonElement("iconSet")]
        [JsonPropertyName("iconSet")]
        public string? IconSet { get; set; } = "3Flags";
    }

    [BsonIgnoreExtraElements]
    public class ExcelChartDataRangeConfig
    {
        [BsonElement("chartSourceFile")]
        [JsonPropertyName("chartSourceFile")]
        public string? ChartSourceFile { get; set; }

        [BsonElement("expectedCategoryRange")]
        [JsonPropertyName("expectedCategoryRange")]
        public string? ExpectedCategoryRange { get; set; }

        [BsonElement("expectedValueRange")]
        [JsonPropertyName("expectedValueRange")]
        public string? ExpectedValueRange { get; set; }

        [BsonElement("expectedPointCount")]
        [JsonPropertyName("expectedPointCount")]
        public int? ExpectedPointCount { get; set; }

        [BsonElement("expectedCategoryText")]
        [JsonPropertyName("expectedCategoryText")]
        public string? ExpectedCategoryText { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class ExcelChartStyleConfig
    {
        [BsonElement("chartSourceFile")]
        [JsonPropertyName("chartSourceFile")]
        public string? ChartSourceFile { get; set; }

        [BsonElement("styleSourceFile")]
        [JsonPropertyName("styleSourceFile")]
        public string? StyleSourceFile { get; set; }

        [BsonElement("styleId")]
        [JsonPropertyName("styleId")]
        public int? StyleId { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class ExcelTableNameConfig
    {
        [BsonElement("worksheetName")]
        [JsonPropertyName("worksheetName")]
        public string? WorksheetName { get; set; }

        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; }

        [BsonElement("expectedName")]
        [JsonPropertyName("expectedName")]
        public string? ExpectedName { get; set; }

        [BsonElement("originalName")]
        [JsonPropertyName("originalName")]
        public string? OriginalName { get; set; }

        [BsonElement("requireOriginalNameAbsent")]
        [JsonPropertyName("requireOriginalNameAbsent")]
        public bool? RequireOriginalNameAbsent { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class ExcelWorksheetPageSetupConfig
    {
        [BsonElement("worksheetName")]
        [JsonPropertyName("worksheetName")]
        public string? WorksheetName { get; set; }

        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; }

        [BsonElement("orientation")]
        [JsonPropertyName("orientation")]
        public string? Orientation { get; set; } = "landscape";
    }

    [BsonIgnoreExtraElements]
    public class ExcelClearCellFormattingConfig
    {
        [BsonElement("worksheetName")]
        [JsonPropertyName("worksheetName")]
        public string? WorksheetName { get; set; }

        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; }

        [BsonElement("range")]
        [JsonPropertyName("range")]
        public string? Range { get; set; }

        [BsonElement("defaultStyleId")]
        [JsonPropertyName("defaultStyleId")]
        public int? DefaultStyleId { get; set; } = 0;
    }

    [BsonIgnoreExtraElements]
    public class ExcelDataModelImportConfig
    {
        [BsonElement("sourceFileName")]
        [JsonPropertyName("sourceFileName")]
        public string? SourceFileName { get; set; }

        [BsonElement("expectedWorksheetName")]
        [JsonPropertyName("expectedWorksheetName")]
        public string? ExpectedWorksheetName { get; set; }

        [BsonElement("expectedConnectionName")]
        [JsonPropertyName("expectedConnectionName")]
        public string? ExpectedConnectionName { get; set; }

        [BsonElement("requireConnection")]
        [JsonPropertyName("requireConnection")]
        public bool? RequireConnection { get; set; } = true;

        [BsonElement("requireDataModel")]
        [JsonPropertyName("requireDataModel")]
        public bool? RequireDataModel { get; set; } = true;

        [BsonElement("requireImportedWorksheet")]
        [JsonPropertyName("requireImportedWorksheet")]
        public bool? RequireImportedWorksheet { get; set; } = true;

        [BsonElement("requireQueryTable")]
        [JsonPropertyName("requireQueryTable")]
        public bool? RequireQueryTable { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class ExcelCompatibilityReportConfig
    {
        [BsonElement("worksheetName")]
        [JsonPropertyName("worksheetName")]
        public string? WorksheetName { get; set; }

        [BsonElement("expectedTexts")]
        [JsonPropertyName("expectedTexts")]
        public List<string> ExpectedTexts { get; set; } = new();

        [BsonElement("requireNewWorksheet")]
        [JsonPropertyName("requireNewWorksheet")]
        public bool? RequireNewWorksheet { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class TextBoxContainsTextConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("expectedText")]
        [JsonPropertyName("expectedText")]
        public string? ExpectedText { get; set; }

        [BsonElement("matchMode")]
        [JsonPropertyName("matchMode")]
        public string? MatchMode { get; set; } = "exact";

        [BsonElement("caseSensitive")]
        [JsonPropertyName("caseSensitive")]
        public bool? CaseSensitive { get; set; } = false;

        [BsonElement("targetOccurrence")]
        [JsonPropertyName("targetOccurrence")]
        public int? TargetOccurrence { get; set; } = 1;

        [BsonElement("requireDefaultPaste")]
        [JsonPropertyName("requireDefaultPaste")]
        public bool? RequireDefaultPaste { get; set; } = true;

        [BsonElement("requireRemovedFromBody")]
        [JsonPropertyName("requireRemovedFromBody")]
        public bool? RequireRemovedFromBody { get; set; } = true;

        [BsonElement("forbiddenTextColors")]
        [JsonPropertyName("forbiddenTextColors")]
        public List<string> ForbiddenTextColors { get; set; } = new();

        [BsonElement("forbiddenRunProperties")]
        [JsonPropertyName("forbiddenRunProperties")]
        public List<string> ForbiddenRunProperties { get; set; } = new();
    }

    [BsonIgnoreExtraElements]
    public class PageMarginsConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("top")]
        [JsonPropertyName("top")]
        public int? Top { get; set; }

        [BsonElement("bottom")]
        [JsonPropertyName("bottom")]
        public int? Bottom { get; set; }

        [BsonElement("left")]
        [JsonPropertyName("left")]
        public int? Left { get; set; }

        [BsonElement("right")]
        [JsonPropertyName("right")]
        public int? Right { get; set; }

        [BsonElement("gutter")]
        [JsonPropertyName("gutter")]
        public int? Gutter { get; set; }

        [BsonElement("requireAllSections")]
        [JsonPropertyName("requireAllSections")]
        public bool? RequireAllSections { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class DocumentStyleSetConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/styles.xml";

        [BsonElement("styleSetName")]
        [JsonPropertyName("styleSetName")]
        public string? StyleSetName { get; set; }

        [BsonElement("expectedFragments")]
        [JsonPropertyName("expectedFragments")]
        public List<string> ExpectedFragments { get; set; } = new();

        [BsonElement("ignoreAttributes")]
        [JsonPropertyName("ignoreAttributes")]
        public List<string> IgnoreAttributes { get; set; } = new();

        [BsonElement("matchPolicy")]
        [JsonPropertyName("matchPolicy")]
        public string? MatchPolicy { get; set; } = XmlGradingMatchPolicies.All;
    }

    [BsonIgnoreExtraElements]
    public class PageBorderConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("requiredStyle")]
        [JsonPropertyName("requiredStyle")]
        public string? RequiredStyle { get; set; } = "single";

        [BsonElement("requiredWidth")]
        [JsonPropertyName("requiredWidth")]
        public int? RequiredWidth { get; set; } = 12;

        [BsonElement("minWidth")]
        [JsonPropertyName("minWidth")]
        public int? MinWidth { get; set; }

        [BsonElement("requiredColor")]
        [JsonPropertyName("requiredColor")]
        public string? RequiredColor { get; set; } = "00B0F0";

        [BsonElement("allowedColors")]
        [JsonPropertyName("allowedColors")]
        public List<string> AllowedColors { get; set; } = new();

        [BsonElement("requireBox")]
        [JsonPropertyName("requireBox")]
        public bool? RequireBox { get; set; } = true;

        [BsonElement("requireAllSections")]
        [JsonPropertyName("requireAllSections")]
        public bool? RequireAllSections { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class SectionBreakBeforeTextConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("targetText")]
        [JsonPropertyName("targetText")]
        public string? TargetText { get; set; }

        [BsonElement("breakType")]
        [JsonPropertyName("breakType")]
        public string? BreakType { get; set; } = "continuous";

        [BsonElement("targetOccurrence")]
        [JsonPropertyName("targetOccurrence")]
        public int? TargetOccurrence { get; set; } = 1;

        [BsonElement("requireImmediateBefore")]
        [JsonPropertyName("requireImmediateBefore")]
        public bool? RequireImmediateBefore { get; set; } = true;

        [BsonElement("allowSameParagraphSectPr")]
        [JsonPropertyName("allowSameParagraphSectPr")]
        public bool? AllowSameParagraphSectPr { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class PictureStyleConfig
    {
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("relsFile")]
        [JsonPropertyName("relsFile")]
        public string? RelsFile { get; set; } = "word/_rels/document.xml.rels";

        [BsonElement("assetId")]
        [JsonPropertyName("assetId")]
        public string? AssetId { get; set; }

        [BsonElement("imageHash")]
        [JsonPropertyName("imageHash")]
        public string? ImageHash { get; set; }

        [BsonElement("perceptualHash")]
        [JsonPropertyName("perceptualHash")]
        public string? PerceptualHash { get; set; }

        [BsonElement("targetImageIndex")]
        [JsonPropertyName("targetImageIndex")]
        public int? TargetImageIndex { get; set; } = 1;

        [BsonElement("stylePreset")]
        [JsonPropertyName("stylePreset")]
        public string? StylePreset { get; set; } = "simpleFrameBlack";

        [BsonElement("requiredLineColor")]
        [JsonPropertyName("requiredLineColor")]
        public string? RequiredLineColor { get; set; } = "000000";

        [BsonElement("minLineWidth")]
        [JsonPropertyName("minLineWidth")]
        public int? MinLineWidth { get; set; }

        [BsonElement("presetGeometry")]
        [JsonPropertyName("presetGeometry")]
        public string? PresetGeometry { get; set; } = "rect";
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

        [BsonElement("anchorTextBefore")]
        [JsonPropertyName("anchorTextBefore")]
        public string? AnchorTextBefore { get; set; }

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
        [BsonElement("sourceFile")]
        [JsonPropertyName("sourceFile")]
        public string? SourceFile { get; set; } = "word/document.xml";

        [BsonElement("relsFile")]
        [JsonPropertyName("relsFile")]
        public string? RelsFile { get; set; } = "word/_rels/document.xml.rels";

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

        [BsonElement("positionConfig")]
        [JsonPropertyName("positionConfig")]
        public ImagePositionConfig? PositionConfig { get; set; }

        [BsonElement("sizeConfig")]
        [JsonPropertyName("sizeConfig")]
        public ImageSizeConfig? SizeConfig { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class ImagePositionConfig
    {
        [BsonElement("afterText")]
        [JsonPropertyName("afterText")]
        public string? AfterText { get; set; }

        [BsonElement("beforeText")]
        [JsonPropertyName("beforeText")]
        public string? BeforeText { get; set; }

        [BsonElement("requireBetween")]
        [JsonPropertyName("requireBetween")]
        public bool? RequireBetween { get; set; } = true;

        [BsonElement("caseSensitive")]
        [JsonPropertyName("caseSensitive")]
        public bool? CaseSensitive { get; set; } = false;
    }

    [BsonIgnoreExtraElements]
    public class ImageSizeConfig
    {
        [BsonElement("expectedWidthEmu")]
        [JsonPropertyName("expectedWidthEmu")]
        public long? ExpectedWidthEmu { get; set; }

        [BsonElement("expectedHeightEmu")]
        [JsonPropertyName("expectedHeightEmu")]
        public long? ExpectedHeightEmu { get; set; }

        [BsonElement("toleranceEmu")]
        [JsonPropertyName("toleranceEmu")]
        public long? ToleranceEmu { get; set; } = 0;
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
