using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project15
{
    public class P15T3Grader : ITaskGrader
    {
        public string TaskId => "P15-T3";
        public string TaskName => "Orders: Conditional Formatting Above Average voi Green style";
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
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Orders'.");
                    return result;
                }

                decimal score = 0m;
                if (!P15GraderHelpers.TryFindOrderTotalDataRange(ws, out var orderTotalCol, out var dataStart, out var dataEnd))
                {
                    result.Errors.Add("Không tìm th?y c?t 'OrderTotal' hoac không xác d?nh duoc vung d? li?u tren sheet Orders.");
                    return result;
                }

                var expectedRange = $"{ExcelCellBase.GetAddress(dataStart, orderTotalCol)}:{ExcelCellBase.GetAddress(dataEnd, orderTotalCol)}";
                var targetedRules = ws.ConditionalFormatting
                    .Where(cf => P15GraderHelpers.SqrefTargetsColumn(cf.Address.Address, orderTotalCol, dataStart, dataEnd))
                    .ToList();

                var matchedRule = targetedRules.FirstOrDefault(cf =>
                    cf.Type == OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.AboveAverage);

                if (matchedRule == null)
                {
                    matchedRule = targetedRules.FirstOrDefault(cf =>
                    {
                        if (cf.Type != OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.Expression
                            && cf.Type != OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.GreaterThan)
                        {
                            return false;
                        }

                        var formula = cf.GetType().GetProperty("Formula")?.GetValue(cf)?.ToString() ?? string.Empty;
                        return P15GraderHelpers.IsFormulaBasedAboveAverageRule(formula, orderTotalCol, dataStart, dataEnd);
                    });
                }

                if (matchedRule == null)
                {
                    if (targetedRules.Count == 0)
                    {
                        result.Errors.Add($"Không tìm th?y conditional formatting rule tren range {expectedRange}.");
                    }
                    else
                    {
                        var types = string.Join(", ", targetedRules.Select(r => r.Type.ToString()).Distinct(StringComparer.OrdinalIgnoreCase));
                        result.Errors.Add($"Co rule tren range {expectedRange} nhung không ph?i quy tac 'Above Average' (types: {types}).");
                    }

                    result.Score = score;
                    return result;
                }

                score += 2m;
                result.Details.Add($"Tìm th?y quy tac Above Average (dòng) tren range {expectedRange}.");

                var fontColor = matchedRule.Style.Font.Color.Color;
                var fillBgColor = matchedRule.Style.Fill.BackgroundColor.Color;
                var fillPatternColor = matchedRule.Style.Fill.PatternColor.Color;

                var fontHex = fontColor.HasValue ? P15GraderHelpers.ToArgbHex(fontColor.Value.ToArgb()) : string.Empty;
                var bgHex = fillBgColor.HasValue ? P15GraderHelpers.ToArgbHex(fillBgColor.Value.ToArgb()) : string.Empty;
                var patternHex = fillPatternColor.HasValue ? P15GraderHelpers.ToArgbHex(fillPatternColor.Value.ToArgb()) : string.Empty;

                var isGreenFill = P15GraderHelpers.IsGreenFillColor(bgHex) || P15GraderHelpers.IsGreenFillColor(patternHex);
                var isDarkGreenText = P15GraderHelpers.IsDarkGreenTextColor(fontHex);
                if (isGreenFill && isDarkGreenText)
                {
                    score += 2m;
                    result.Details.Add("Style quy tac dung yêu c?u: Green Fill with Dark Green Text.");
                }
                else
                {
                    result.Errors.Add($"Style quy tac chua dúng Green Fill with Dark Green Text. fill.bg={bgHex}, fill.pattern={patternHex}, font={fontHex}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




