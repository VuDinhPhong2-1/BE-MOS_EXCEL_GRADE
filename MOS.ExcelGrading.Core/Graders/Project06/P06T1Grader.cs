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
                    result.Errors.Add("Không tìm thấy sheet 'Summary'.");
                    return result;
                }

                var targetRule = ws.ConditionalFormatting.FirstOrDefault(cf =>
                    cf.Type == OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.GreaterThan &&
                    P06GraderHelpers.NormalizeAddress(cf.Address.Address) == "F4:F11");
                if (targetRule == null)
                {
                    result.Errors.Add("Không tìm thấy rule Conditional Formatting đúng cho F4:F11 với điều kiện Greater Than.");
                    return result;
                }

                decimal score = 1m; // Tim thay rule dung range + operator.
                result.Details.Add("Đã tìm thấy rule Conditional Formatting tại F4:F11.");

                var formula = targetRule.GetType().GetProperty("Formula")?.GetValue(targetRule)?.ToString() ?? string.Empty;
                var formulaNormalized = P06GraderHelpers.NormalizeFormula(formula);
                if (formulaNormalized == "5000000")
                {
                    score += 1m;
                    result.Details.Add("Điều kiện so sánh đúng ngưỡng 5,000,000.");
                }
                else
                {
                    result.Errors.Add($"Ngưỡng Greater Than chưa đúng. Hiện tại: '{formula}'.");
                }

                var fontColor = targetRule.Style.Font.Color.Color;
                var bgColor = targetRule.Style.Fill.BackgroundColor.Color;
                if (!fontColor.HasValue || !bgColor.HasValue)
                {
                    result.Errors.Add("Không đọc được style màu của Conditional Formatting.");
                    result.Score = Math.Min(MaxScore, score);
                    return result;
                }

                var fontArgb = fontColor.Value.ToArgb();
                var bgArgb = bgColor.Value.ToArgb();

                if (fontArgb == unchecked((int)0xFF9C5700))
                {
                    score += 1m;
                    result.Details.Add("Màu chữ dùng Dark Yellow (FF9C5700).");
                }
                else
                {
                    result.Errors.Add($"Màu chữ chưa đúng Dark Yellow. Hiện tại: ARGB {fontArgb:X8}.");
                }

                if (bgArgb == unchecked((int)0xFFFFEB9C))
                {
                    score += 1m;
                    result.Details.Add("Màu nền dùng Yellow (FFFFEB9C).");
                }
                else
                {
                    result.Errors.Add($"Màu nền chưa đúng Yellow. Hiện tại: ARGB {bgArgb:X8}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}

// minor-sync: non-functional graders update
