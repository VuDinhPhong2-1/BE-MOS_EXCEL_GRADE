using MongoDB.Bson.Serialization.Attributes;
using System.Text.Json.Serialization;

namespace MOS.ExcelGrading.Core.Models
{
    public partial class SpecialCondition
    {
        [BsonElement("pptPictureCropShapeConfig")]
        [JsonPropertyName("pptPictureCropShapeConfig")]
        [BsonIgnoreIfNull]
        public PptPictureCropShapeConfig? PptPictureCropShapeConfig { get; set; }

        [BsonElement("pptShapeSizeConfig")]
        [JsonPropertyName("pptShapeSizeConfig")]
        [BsonIgnoreIfNull]
        public PptShapeSizeConfig? PptShapeSizeConfig { get; set; }

        [BsonElement("pptShapeGroupConfig")]
        [JsonPropertyName("pptShapeGroupConfig")]
        [BsonIgnoreIfNull]
        public PptShapeGroupConfig? PptShapeGroupConfig { get; set; }

        [BsonElement("pptChartLegendConfig")]
        [JsonPropertyName("pptChartLegendConfig")]
        [BsonIgnoreIfNull]
        public PptChartLegendConfig? PptChartLegendConfig { get; set; }

        [BsonElement("pptSmartArtConfig")]
        [JsonPropertyName("pptSmartArtConfig")]
        [BsonIgnoreIfNull]
        public PptSmartArtConfig? PptSmartArtConfig { get; set; }

        [BsonElement("pptCommentConfig")]
        [JsonPropertyName("pptCommentConfig")]
        [BsonIgnoreIfNull]
        public PptCommentConfig? PptCommentConfig { get; set; }

        [BsonElement("pptSlideTitlesConfig")]
        [JsonPropertyName("pptSlideTitlesConfig")]
        [BsonIgnoreIfNull]
        public PptSlideTitlesConfig? PptSlideTitlesConfig { get; set; }

        [BsonElement("pptVideoConfig")]
        [JsonPropertyName("pptVideoConfig")]
        [BsonIgnoreIfNull]
        public PptVideoConfig? PptVideoConfig { get; set; }

        [BsonElement("pptTableConfig")]
        [JsonPropertyName("pptTableConfig")]
        [BsonIgnoreIfNull]
        public PptTableConfig? PptTableConfig { get; set; }

        [BsonElement("pptSectionConfig")]
        [JsonPropertyName("pptSectionConfig")]
        [BsonIgnoreIfNull]
        public PptSectionConfig? PptSectionConfig { get; set; }

        [BsonElement("pptPictureStyleConfig")]
        [JsonPropertyName("pptPictureStyleConfig")]
        [BsonIgnoreIfNull]
        public PptPictureStyleConfig? PptPictureStyleConfig { get; set; }

        [BsonElement("pptShapeArrangeConfig")]
        [JsonPropertyName("pptShapeArrangeConfig")]
        [BsonIgnoreIfNull]
        public PptShapeArrangeConfig? PptShapeArrangeConfig { get; set; }

        [BsonElement("pptSummaryZoomConfig")]
        [JsonPropertyName("pptSummaryZoomConfig")]
        [BsonIgnoreIfNull]
        public PptSummaryZoomConfig? PptSummaryZoomConfig { get; set; }

        [BsonElement("pptExportedFileConfig")]
        [JsonPropertyName("pptExportedFileConfig")]
        [BsonIgnoreIfNull]
        public PptExportedFileConfig? PptExportedFileConfig { get; set; }

        [BsonElement("pptMasterPictureConfig")]
        [JsonPropertyName("pptMasterPictureConfig")]
        [BsonIgnoreIfNull]
        public PptMasterPictureConfig? PptMasterPictureConfig { get; set; }

        [BsonElement("pptSlideTransitionConfig")]
        [JsonPropertyName("pptSlideTransitionConfig")]
        [BsonIgnoreIfNull]
        public PptSlideTransitionConfig? PptSlideTransitionConfig { get; set; }

        [BsonElement("pptAnimationConfig")]
        [JsonPropertyName("pptAnimationConfig")]
        [BsonIgnoreIfNull]
        public PptAnimationConfig? PptAnimationConfig { get; set; }

        [BsonElement("pptMarkAsFinalConfig")]
        [JsonPropertyName("pptMarkAsFinalConfig")]
        [BsonIgnoreIfNull]
        public PptMarkAsFinalConfig? PptMarkAsFinalConfig { get; set; }

        [BsonElement("pptPrintSettingsConfig")]
        [JsonPropertyName("pptPrintSettingsConfig")]
        [BsonIgnoreIfNull]
        public PptPrintSettingsConfig? PptPrintSettingsConfig { get; set; }

        [BsonElement("pptTextColumnsConfig")]
        [JsonPropertyName("pptTextColumnsConfig")]
        [BsonIgnoreIfNull]
        public PptTextColumnsConfig? PptTextColumnsConfig { get; set; }

        [BsonElement("pptNotesMasterPlaceholdersConfig")]
        [JsonPropertyName("pptNotesMasterPlaceholdersConfig")]
        [BsonIgnoreIfNull]
        public PptNotesMasterPlaceholdersConfig? PptNotesMasterPlaceholdersConfig { get; set; }

        [BsonElement("pptSlideSizeConfig")]
        [JsonPropertyName("pptSlideSizeConfig")]
        [BsonIgnoreIfNull]
        public PptSlideSizeConfig? PptSlideSizeConfig { get; set; }

        [BsonElement("pptChartTypeConfig")]
        [JsonPropertyName("pptChartTypeConfig")]
        [BsonIgnoreIfNull]
        public PptChartTypeConfig? PptChartTypeConfig { get; set; }

        [BsonElement("pptAltTextConfig")]
        [JsonPropertyName("pptAltTextConfig")]
        [BsonIgnoreIfNull]
        public PptAltTextConfig? PptAltTextConfig { get; set; }

        [BsonElement("pptHyperlinkConfig")]
        [JsonPropertyName("pptHyperlinkConfig")]
        [BsonIgnoreIfNull]
        public PptHyperlinkConfig? PptHyperlinkConfig { get; set; }

        [BsonElement("pptTextBoxConfig")]
        [JsonPropertyName("pptTextBoxConfig")]
        [BsonIgnoreIfNull]
        public PptTextBoxConfig? PptTextBoxConfig { get; set; }

        [BsonElement("pptSlideLayoutConfig")]
        [JsonPropertyName("pptSlideLayoutConfig")]
        [BsonIgnoreIfNull]
        public PptSlideLayoutConfig? PptSlideLayoutConfig { get; set; }

        [BsonElement("pptShapeStyleConfig")]
        [JsonPropertyName("pptShapeStyleConfig")]
        [BsonIgnoreIfNull]
        public PptShapeStyleConfig? PptShapeStyleConfig { get; set; }

        [BsonElement("pptSlideBackgroundConfig")]
        [JsonPropertyName("pptSlideBackgroundConfig")]
        [BsonIgnoreIfNull]
        public PptSlideBackgroundConfig? PptSlideBackgroundConfig { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptSlideRef
    {
        [BsonElement("slideIndex")]
        [JsonPropertyName("slideIndex")]
        public int? SlideIndex { get; set; }

        [BsonElement("slideTitle")]
        [JsonPropertyName("slideTitle")]
        public string? SlideTitle { get; set; }

        [BsonElement("anchorText")]
        [JsonPropertyName("anchorText")]
        public string? AnchorText { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptShapeSelector
    {
        [BsonElement("shapeName")]
        [JsonPropertyName("shapeName")]
        public string? ShapeName { get; set; }

        [BsonElement("targetText")]
        [JsonPropertyName("targetText")]
        public string? TargetText { get; set; }

        [BsonElement("shapeType")]
        [JsonPropertyName("shapeType")]
        public string? ShapeType { get; set; }

        [BsonElement("ordinal")]
        [JsonPropertyName("ordinal")]
        public int? Ordinal { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptPictureCropShapeConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("expectedShapePreset")]
        [JsonPropertyName("expectedShapePreset")]
        public string ExpectedShapePreset { get; set; } = "ellipse"; // Oval / ellipse
    }

    [BsonIgnoreExtraElements]
    public class PptShapeSizeConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("expectedHeightInches")]
        [JsonPropertyName("expectedHeightInches")]
        public double? ExpectedHeightInches { get; set; }

        [BsonElement("expectedWidthInches")]
        [JsonPropertyName("expectedWidthInches")]
        public double? ExpectedWidthInches { get; set; }

        [BsonElement("toleranceInches")]
        [JsonPropertyName("toleranceInches")]
        public double? ToleranceInches { get; set; } = 0.15;
    }

    [BsonIgnoreExtraElements]
    public class PptShapeGroupConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("requireGrouped")]
        [JsonPropertyName("requireGrouped")]
        public bool RequireGrouped { get; set; } = true;

        [BsonElement("minChildCount")]
        [JsonPropertyName("minChildCount")]
        public int? MinChildCount { get; set; }

        [BsonElement("checkAlignCenter")]
        [JsonPropertyName("checkAlignCenter")]
        public bool? CheckAlignCenter { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptChartLegendConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("expectedPosition")]
        [JsonPropertyName("expectedPosition")]
        public string ExpectedPosition { get; set; } = "t"; // t, b, l, r, tr
    }

    [BsonIgnoreExtraElements]
    public class PptSmartArtConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("expectedLayoutName")]
        [JsonPropertyName("expectedLayoutName")]
        public string? ExpectedLayoutName { get; set; } // Target List, Upward Arrow

        [BsonElement("expectedLayoutIdEnding")]
        [JsonPropertyName("expectedLayoutIdEnding")]
        public string? ExpectedLayoutIdEnding { get; set; }

        [BsonElement("expectedNodeTexts")]
        [JsonPropertyName("expectedNodeTexts")]
        public List<string> ExpectedNodeTexts { get; set; } = new();

        [BsonElement("requireOrderedNodes")]
        [JsonPropertyName("requireOrderedNodes")]
        public bool? RequireOrderedNodes { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptCommentConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("author")]
        [JsonPropertyName("author")]
        public string? Author { get; set; }

        [BsonElement("targetCommentText")]
        [JsonPropertyName("targetCommentText")]
        public string? TargetCommentText { get; set; }

        [BsonElement("expectAbsent")]
        [JsonPropertyName("expectAbsent")]
        public bool ExpectAbsent { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class PptSlideTitlesConfig
    {
        [BsonElement("expectedTitlesInOrder")]
        [JsonPropertyName("expectedTitlesInOrder")]
        public List<string> ExpectedTitlesInOrder { get; set; } = new();

        [BsonElement("startSlideIndex")]
        [JsonPropertyName("startSlideIndex")]
        public int? StartSlideIndex { get; set; } = 1;
    }

    [BsonIgnoreExtraElements]
    public class PptVideoConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("requireVideoOnly")]
        [JsonPropertyName("requireVideoOnly")]
        public bool RequireVideoOnly { get; set; } = true;

        [BsonElement("expectedTrimEndMs")]
        [JsonPropertyName("expectedTrimEndMs")]
        public long? ExpectedTrimEndMs { get; set; }

        [BsonElement("expectedTrimEndSeconds")]
        [JsonPropertyName("expectedTrimEndSeconds")]
        public double? ExpectedTrimEndSeconds { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptTableConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("disallowedRowTexts")]
        [JsonPropertyName("disallowedRowTexts")]
        public List<string> DisallowedRowTexts { get; set; } = new();

        [BsonElement("requiredCellTexts")]
        [JsonPropertyName("requiredCellTexts")]
        public List<string> RequiredCellTexts { get; set; } = new();

        [BsonElement("expectedRowCount")]
        [JsonPropertyName("expectedRowCount")]
        public int? ExpectedRowCount { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptSectionConfig
    {
        [BsonElement("expectedSectionName")]
        [JsonPropertyName("expectedSectionName")]
        public string ExpectedSectionName { get; set; } = string.Empty;

        [BsonElement("disallowedSectionNames")]
        [JsonPropertyName("disallowedSectionNames")]
        public List<string> DisallowedSectionNames { get; set; } = new();
    }

    [BsonIgnoreExtraElements]
    public class PptPictureStyleConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("expectedStyleName")]
        [JsonPropertyName("expectedStyleName")]
        public string ExpectedStyleName { get; set; } = "Bevel Rectangle";
    }

    [BsonIgnoreExtraElements]
    public class PptShapeArrangeConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("checkAlignMiddle")]
        [JsonPropertyName("checkAlignMiddle")]
        public bool? CheckAlignMiddle { get; set; }

        [BsonElement("checkBringToFront")]
        [JsonPropertyName("checkBringToFront")]
        public bool? CheckBringToFront { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptSummaryZoomConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("requireZoomItems")]
        [JsonPropertyName("requireZoomItems")]
        public bool RequireZoomItems { get; set; } = true;

        [BsonElement("excludeFirstSlide")]
        [JsonPropertyName("excludeFirstSlide")]
        public bool? ExcludeFirstSlide { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class PptExportedFileConfig
    {
        [BsonElement("expectedFileName")]
        [JsonPropertyName("expectedFileName")]
        public string ExpectedFileName { get; set; } = "Backcountry.pdf";

        [BsonElement("caseSensitive")]
        [JsonPropertyName("caseSensitive")]
        public bool CaseSensitive { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class PptMasterPictureConfig
    {
        [BsonElement("expectedPlacement")]
        [JsonPropertyName("expectedPlacement")]
        public string? ExpectedPlacement { get; set; } = "bottom-right";

        [BsonElement("expectedImageName")]
        [JsonPropertyName("expectedImageName")]
        public string? ExpectedImageName { get; set; } // Badge.png
    }

    [BsonIgnoreExtraElements]
    public class PptSlideTransitionConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("expectedTransition")]
        [JsonPropertyName("expectedTransition")]
        public string? ExpectedTransition { get; set; } // Smoothly / Fade

        [BsonElement("expectedDurationSeconds")]
        [JsonPropertyName("expectedDurationSeconds")]
        public double? ExpectedDurationSeconds { get; set; } // 0.75

        [BsonElement("applyToAll")]
        [JsonPropertyName("applyToAll")]
        public bool? ApplyToAll { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class PptAnimationConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("expectedEffect")]
        [JsonPropertyName("expectedEffect")]
        public string? ExpectedEffect { get; set; } // Fly In, Motion Path Square

        [BsonElement("motionPathType")]
        [JsonPropertyName("motionPathType")]
        public string? MotionPathType { get; set; } // Square
    }

    [BsonIgnoreExtraElements]
    public class PptMarkAsFinalConfig
    {
        [BsonElement("expectedMarkAsFinal")]
        [JsonPropertyName("expectedMarkAsFinal")]
        public bool ExpectedMarkAsFinal { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class PptPrintSettingsConfig
    {
        [BsonElement("expectedPrintWhat")]
        [JsonPropertyName("expectedPrintWhat")]
        public string? ExpectedPrintWhat { get; set; } // slides, notes

        [BsonElement("expectedColorMode")]
        [JsonPropertyName("expectedColorMode")]
        public string? ExpectedColorMode { get; set; } // gray, color
    }

    [BsonIgnoreExtraElements]
    public class PptTextColumnsConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("expectedColumnCount")]
        [JsonPropertyName("expectedColumnCount")]
        public int ExpectedColumnCount { get; set; } = 2;

        [BsonElement("expectedSpacingInches")]
        [JsonPropertyName("expectedSpacingInches")]
        public double? ExpectedSpacingInches { get; set; } = 0.5;
    }

    [BsonIgnoreExtraElements]
    public class PptNotesMasterPlaceholdersConfig
    {
        [BsonElement("requireHeaderDisabled")]
        [JsonPropertyName("requireHeaderDisabled")]
        public bool RequireHeaderDisabled { get; set; } = true;

        [BsonElement("requireFooterDisabled")]
        [JsonPropertyName("requireFooterDisabled")]
        public bool RequireFooterDisabled { get; set; } = true;
    }

    [BsonIgnoreExtraElements]
    public class PptSlideSizeConfig
    {
        [BsonElement("expectedRatio")]
        [JsonPropertyName("expectedRatio")]
        public string ExpectedRatio { get; set; } = "16:9"; // 16:9 / widescreen
    }

    [BsonIgnoreExtraElements]
    public class PptChartTypeConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("expectedChartType")]
        [JsonPropertyName("expectedChartType")]
        public string ExpectedChartType { get; set; } = "Pareto";
    }

    [BsonIgnoreExtraElements]
    public class PptAltTextConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("expectedAltText")]
        [JsonPropertyName("expectedAltText")]
        public string ExpectedAltText { get; set; } = string.Empty;
    }

    [BsonIgnoreExtraElements]
    public class PptHyperlinkConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("targetText")]
        [JsonPropertyName("targetText")]
        public string? TargetText { get; set; }

        [BsonElement("expectedUrl")]
        [JsonPropertyName("expectedUrl")]
        public string ExpectedUrl { get; set; } = string.Empty;
    }

    [BsonIgnoreExtraElements]
    public class PptTextBoxConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("expectedText")]
        [JsonPropertyName("expectedText")]
        public string ExpectedText { get; set; } = string.Empty;

        [BsonElement("expectedWidthInches")]
        [JsonPropertyName("expectedWidthInches")]
        public double? ExpectedWidthInches { get; set; } = 2.5;

        [BsonElement("expectedPlacement")]
        [JsonPropertyName("expectedPlacement")]
        public string? ExpectedPlacement { get; set; } = "bottom-right";
    }

    [BsonIgnoreExtraElements]
    public class PptSlideLayoutConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("expectedLayoutName")]
        [JsonPropertyName("expectedLayoutName")]
        public string ExpectedLayoutName { get; set; } = "Two Content";
    }

    [BsonIgnoreExtraElements]
    public class PptShapeStyleConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("shape")]
        [JsonPropertyName("shape")]
        public PptShapeSelector? Shape { get; set; }

        [BsonElement("expectedStyleName")]
        [JsonPropertyName("expectedStyleName")]
        public string ExpectedStyleName { get; set; } = "Intense Effect - Blue-Gray, Accent 1";
    }

    [BsonIgnoreExtraElements]
    public class PptSlideBackgroundConfig
    {
        [BsonElement("slide")]
        [JsonPropertyName("slide")]
        public PptSlideRef? Slide { get; set; }

        [BsonElement("expectedColorHex")]
        [JsonPropertyName("expectedColorHex")]
        public string ExpectedColorHex { get; set; } = "0070C0"; // Blue
    }
}
