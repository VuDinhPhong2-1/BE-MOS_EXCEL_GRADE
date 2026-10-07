using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public partial class XmlGradingRuleService
    {
        private static bool TryValidatePptSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var type = specialCondition.Type;
            if (!IsPptSpecialConditionSupported(type)) return false;

            if (string.Equals(type, SpecialConditionTypes.PptPictureCropShape, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptPictureCropShapeConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptPictureCropShapeConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptShapeSize, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptShapeSizeConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptShapeSizeConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptShapeGroup, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptShapeGroupConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptShapeGroupConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptChartLegend, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptChartLegendConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptChartLegendConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSmartArt, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSmartArtConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSmartArtConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptComment, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptCommentConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptCommentConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSlideTitles, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSlideTitlesConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSlideTitlesConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptVideo, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptVideoConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptVideoConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptTable, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptTableConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptTableConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSection, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSectionConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSectionConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptPictureStyle, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptPictureStyleConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptPictureStyleConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptShapeArrange, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptShapeArrangeConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptShapeArrangeConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSummaryZoom, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSummaryZoomConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSummaryZoomConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptExportedFile, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptExportedFileConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptExportedFileConfig không được null.");
                else if (string.IsNullOrWhiteSpace(specialCondition.PptExportedFileConfig.ExpectedFileName))
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptExportedFileConfig.expectedFileName không được rỗng.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptMasterPicture, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptMasterPictureConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptMasterPictureConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSlideTransition, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSlideTransitionConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSlideTransitionConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptAnimation, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptAnimationConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptAnimationConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptMarkAsFinal, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptMarkAsFinalConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptMarkAsFinalConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptPrintSettings, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptPrintSettingsConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptPrintSettingsConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptTextColumns, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptTextColumnsConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptTextColumnsConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptNotesMasterPlaceholders, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptNotesMasterPlaceholdersConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptNotesMasterPlaceholdersConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSlideSize, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSlideSizeConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSlideSizeConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptChartType, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptChartTypeConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptChartTypeConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptAltText, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptAltTextConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptAltTextConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptHyperlink, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptHyperlinkConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptHyperlinkConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptTextBox, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptTextBoxConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptTextBoxConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSlideLayout, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSlideLayoutConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSlideLayoutConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptShapeStyle, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptShapeStyleConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptShapeStyleConfig không được null.");
                return true;
            }

            if (string.Equals(type, SpecialConditionTypes.PptSlideBackground, StringComparison.OrdinalIgnoreCase))
            {
                if (specialCondition.PptSlideBackgroundConfig == null)
                    result.Errors.Add($"{taskPrefix}.specialCondition.pptSlideBackgroundConfig không được null.");
                return true;
            }

            return false;
        }
    }
}
