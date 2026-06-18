using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project06
{
    public class P06T1Grader : ITaskGrader
    {
        public string TaskId => "P06-T1";
        public string TaskName => "Conditional Formatting F4:F11 > 5,000,000 (Yellow fill + Dark Yellow text)";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet)
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
                    TaskResultIssueHelper.AddIssue(result, "Khong tim thay sheet 'Summary'.");
                    return result;
                }

                var targetRule = ws.ConditionalFormatting.FirstOrDefault(cf =>
                    cf.Type == OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.GreaterThan &&
                    P06GraderHelpers.NormalizeAddress(cf.Address.Address) == "F4:F11");

                if (targetRule == null)
                {
                    TaskResultIssueHelper.AddIssue(
                        result,
                        "Khong tim thay rule Conditional Formatting dung cho F4:F11 voi dieu kien Greater Than.",
                        "Tao Conditional Formatting cho vung F4:F11 voi quy tac Greater Than 5000000.");
                    return result;
                }

                decimal score = 1m;
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
                    TaskResultIssueHelper.AddIssue(
                        result,
                        $"Nguong Greater Than chua dung. Hien tai: '{formula}'.",
                        "Sua lai Conditional Formatting rule thanh Greater Than 5000000 cho vung F4:F11.");
                }

                var fontColor = targetRule.Style.Font.Color.Color;
                var bgColor = targetRule.Style.Fill.BackgroundColor.Color;
                if (!fontColor.HasValue || !bgColor.HasValue)
                {
                    TaskResultIssueHelper.AddIssue(
                        result,
                        "Khong doc duoc style mau cua Conditional Formatting.",
                        "Mo lai rule Conditional Formatting va chon dung mau chu Dark Yellow cung nen Yellow.");
                    result.Score = Math.Min(MaxScore, score);
                    return result;
                }

                var fontArgb = fontColor.Value.ToArgb();
                var bgArgb = bgColor.Value.ToArgb();

                if (fontArgb == unchecked((int)0xFF9C5700))
                {
                    score += 1m;
                    result.Details.Add("Mau chu dung Dark Yellow (FF9C5700).");
                }
                else
                {
                    TaskResultIssueHelper.AddIssue(
                        result,
                        $"Mau chu chua dung Dark Yellow. Hien tai: ARGB {fontArgb:X8}.",
                        "Trong Conditional Formatting, doi Font Color thanh Dark Yellow.");
                }

                if (bgArgb == unchecked((int)0xFFFFEB9C))
                {
                    score += 1m;
                    result.Details.Add("Mau nen dung Yellow (FFFFEB9C).");
                }
                else
                {
                    TaskResultIssueHelper.AddIssue(
                        result,
                        $"Mau nen chua dung Yellow. Hien tai: ARGB {bgArgb:X8}.",
                        "Trong Conditional Formatting, doi Fill/Background thanh Yellow.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                TaskResultIssueHelper.AddIssue(result, $"Loi: {ex.Message}");
            }

            return result;
        }
    }
}

