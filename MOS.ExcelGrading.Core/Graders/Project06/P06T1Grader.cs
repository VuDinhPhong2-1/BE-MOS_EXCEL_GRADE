using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project06
{
    public class P06T1Grader : ITaskGrader
    {
        public string TaskId => "P06-T1";
        public string TaskName => "Conditional Formatting F4:F11 > 5,000,000 (Yellow fill + Dark Yellow text)";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P06GraderHelpers.GetSheet(studentSheet, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Summary'.");
                    return result;
                }

                var targetRule = ws.ConditionalFormatting.FirstOrDefault(cf =>
                    cf.Type == OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.GreaterThan &&
                    P06GraderHelpers.NormalizeAddress(cf.Address.Address) == "F4:F11");
                if (targetRule == null)
                {
                    result.Errors.Add("Khong tim thay rule Conditional Formatting dung cho F4:F11 voi dieu kien Greater Than.");
                    return result;
                }

                decimal score = 1m; // Tim thay rule dung range + operator.
                result.Details.Add("Da tim thay rule Conditional Formatting tai F4:F11.");

                var formula = targetRule.GetType().GetProperty("Formula")?.GetValue(targetRule)?.ToString() ?? string.Empty;
                var formulaNormalized = P06GraderHelpers.NormalizeFormula(formula);
                if (formulaNormalized == "5000000")
                {
                    score += 1m;
                    result.Details.Add("Dieu kien so sanh dung nguong 5,000,000.");
                }
                else
                {
                    result.Errors.Add($"Nguong Greater Than chua dung. Hien tai: '{formula}'.");
                }

                var fontColor = targetRule.Style.Font.Color.Color;
                var bgColor = targetRule.Style.Fill.BackgroundColor.Color;
                if (!fontColor.HasValue || !bgColor.HasValue)
                {
                    result.Errors.Add("Khong doc duoc style mau cua Conditional Formatting.");
                    result.Score = Math.Min(MaxScore, score);
                    return result;
                }

                var fontArgb = fontColor.Value.ToArgb();
                var bgArgb = bgColor.Value.ToArgb();

                if (fontArgb == unchecked((int)0xFF9C5700))
                {
                    score += 1m;
                    result.Details.Add("Mau chu dxf dung Dark Yellow (FF9C5700).");
                }
                else
                {
                    result.Errors.Add($"Mau chu chua dung Dark Yellow. Hien tai: ARGB {fontArgb:X8}.");
                }

                if (bgArgb == unchecked((int)0xFFFFEB9C))
                {
                    score += 1m;
                    result.Details.Add("Mau nen dxf dung Yellow (FFFFEB9C).");
                }
                else
                {
                    result.Errors.Add($"Mau nen chua dung Yellow. Hien tai: ARGB {bgArgb:X8}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
