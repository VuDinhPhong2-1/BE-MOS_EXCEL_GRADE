using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project16
{
    public class P16T3Grader : ITaskGrader
    {
        public string TaskId => "P16-T3";
        public string TaskName => "Products: ap dung Icon Set 3 den giao thong cho Quantity";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Products'.");
                    return result;
                }

                decimal score = 0m;
                var quantityCol = -1;
                var dataStart = -1;
                var dataEnd = -1;

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => P16GraderHelpers.IsQuantityColumnName(c.Name)));
                if (table != null)
                {
                    var quantityOffset = table.Columns
                        .Select((c, idx) => new { Column = c, Index = idx })
                        .First(x => P16GraderHelpers.IsQuantityColumnName(x.Column.Name))
                        .Index;
                    quantityCol = table.Address.Start.Column + quantityOffset;
                    dataStart = table.Address.Start.Row + 1;
                    dataEnd = table.Address.End.Row;
                }
                else if (P16GraderHelpers.TryFindColumnByHeader(ws, P16GraderHelpers.IsQuantityColumnName, out var headerRow, out var headerCol))
                {
                    quantityCol = headerCol;
                    dataStart = headerRow + 1;
                    dataEnd = P16GraderHelpers.GetLastDataRowInColumn(ws, headerCol, dataStart);
                }
                else
                {
                    result.Errors.Add("Không tìm th?y c?t 'Quantity'.");
                    return result;
                }

                if (dataStart <= 0 || dataEnd < dataStart)
                {
                    result.Errors.Add("Không xác d?nh du?c vung d? li?u c?t Quantity.");
                    return result;
                }

                var expectedRange = $"{ExcelCellBase.GetAddress(dataStart, quantityCol)}:{ExcelCellBase.GetAddress(dataEnd, quantityCol)}";
                var targetedRules = ws.ConditionalFormatting
                    .Where(cf => P16GraderHelpers.SqrefTargetsColumn(cf.Address?.Address, quantityCol, dataStart, dataEnd))
                    .ToList();

                var matchedRule = targetedRules.FirstOrDefault(cf =>
                    cf.Type == OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.ThreeIconSet);

                if (matchedRule == null)
                {
                    if (targetedRules.Count == 0)
                    {
                        result.Errors.Add($"Không tìm th?y icon set tren range {expectedRange}.");
                    }
                    else
                    {
                        var types = string.Join(", ", targetedRules.Select(r => r.Type.ToString()).Distinct(StringComparer.OrdinalIgnoreCase));
                        result.Errors.Add($"Co conditional formatting tren range {expectedRange} nhung không ph?i 3-icon set (types: {types}).");
                    }

                    result.Score = score;
                    return result;
                }

                score += 2m;
                result.Details.Add($"Tìm th?y icon set tren range {expectedRange}.");

                static bool TryReadIconThreshold(object? iconNode, out string type, out decimal value)
                {
                    type = string.Empty;
                    value = 0m;
                    if (iconNode == null)
                    {
                        return false;
                    }

                    type = iconNode.GetType().GetProperty("Type")?.GetValue(iconNode)?.ToString() ?? string.Empty;
                    var valueText = iconNode.GetType().GetProperty("Value")?.GetValue(iconNode)?.ToString() ?? string.Empty;
                    return decimal.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
                }

                var iconSetName = matchedRule.GetType().GetProperty("IconSet")?.GetValue(matchedRule)?.ToString() ?? string.Empty;
                var validIconSet = string.IsNullOrWhiteSpace(iconSetName)
                                   || iconSetName.Contains("TrafficLights1", StringComparison.OrdinalIgnoreCase)
                                   || iconSetName.Contains("TrafficLights2", StringComparison.OrdinalIgnoreCase)
                                   || string.Equals(iconSetName, "3TrafficLights1", StringComparison.OrdinalIgnoreCase)
                                   || string.Equals(iconSetName, "3TrafficLights2", StringComparison.OrdinalIgnoreCase);

                var icon1 = matchedRule.GetType().GetProperty("Icon1")?.GetValue(matchedRule);
                var icon2 = matchedRule.GetType().GetProperty("Icon2")?.GetValue(matchedRule);
                var icon3 = matchedRule.GetType().GetProperty("Icon3")?.GetValue(matchedRule);

                var icon1Ok = TryReadIconThreshold(icon1, out var type1, out var val1);
                var icon2Ok = TryReadIconThreshold(icon2, out var type2, out var val2);
                var icon3Ok = TryReadIconThreshold(icon3, out var type3, out var val3);

                var valuesOk = icon1Ok
                               && icon2Ok
                               && icon3Ok
                               && string.Equals(type1, "Percent", StringComparison.OrdinalIgnoreCase)
                               && string.Equals(type2, "Percent", StringComparison.OrdinalIgnoreCase)
                               && string.Equals(type3, "Percent", StringComparison.OrdinalIgnoreCase)
                               && val1 == 0m
                               && val2 == 33m
                               && val3 == 67m;

                if (validIconSet && valuesOk)
                {
                    score += 2m;
                    result.Details.Add("Icon set dung nguong 0/33/67 (3 traffic lights).");
                }
                else
                {
                    result.Errors.Add($"Icon set/nguong chua dúng. iconSet='{iconSetName}', icon1={type1}:{val1}, icon2={type2}:{val2}, icon3={type3}:{val3}.");
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




