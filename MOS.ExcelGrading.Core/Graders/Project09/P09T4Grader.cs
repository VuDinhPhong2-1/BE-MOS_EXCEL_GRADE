using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    public class P09T4Grader : ITaskGrader
    {
        public string TaskId => "P09-T4";
        public string TaskName => "Filter Total column: 34,000 to 45,000";
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
                var ws = P09GraderHelpers.GetSheet(studentSheet.Workbook, "Data");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Data'.");
                    return result;
                }

                decimal score = 0m;

                var filterAddress = ws.AutoFilter?.Address?.Address ?? string.Empty;
                if (string.Equals(
                        P09GraderHelpers.NormalizeRange(filterAddress),
                        "A17:K29",
                        StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("AutoFilter dung vung A17:K29.");
                }
                else
                {
                    result.Errors.Add($"AutoFilter chua dung. Hien tai: '{filterAddress}'.");
                }

                var ns = P09GraderHelpers.CreateWorksheetNamespaceManager(ws.WorksheetXml);
                var autoFilterNode = ws.WorksheetXml.SelectSingleNode("//x:autoFilter", ns);
                if (autoFilterNode == null)
                {
                    result.Errors.Add("Khong tim thay node autoFilter trong XML.");
                    result.Score = score;
                    return result;
                }

                var filterColumnNode = autoFilterNode.SelectSingleNode("x:filterColumn[@colId='10']", ns);
                if (filterColumnNode != null)
                {
                    score += 1m;
                    result.Details.Add("Filter dat dung cot Total (colId=10).");
                }
                else
                {
                    result.Errors.Add("Khong tim thay filterColumn voi colId=10.");
                }

                var customFiltersNode = filterColumnNode?.SelectSingleNode("x:customFilters", ns);
                var andAttr = customFiltersNode?.Attributes?["and"]?.Value ?? string.Empty;
                var hasAnd = string.Equals(andAttr, "1", StringComparison.OrdinalIgnoreCase)
                             || string.Equals(andAttr, "true", StringComparison.OrdinalIgnoreCase);

                if (!hasAnd)
                {
                    result.Errors.Add("customFilters chua dat dieu kien AND.");
                }

                var hasLowerBound = HasCustomFilter(customFiltersNode, ns, "greaterThanOrEqual", 34000m);
                var hasUpperBound = HasCustomFilter(customFiltersNode, ns, "lessThanOrEqual", 45000m);
                if (hasAnd && hasLowerBound && hasUpperBound)
                {
                    score += 2m;
                    result.Details.Add("Dieu kien filter dung: >= 34000 va <= 45000.");
                }
                else
                {
                    if (!hasLowerBound)
                    {
                        result.Errors.Add("Thieu dieu kien lower bound: >= 34000.");
                    }

                    if (!hasUpperBound)
                    {
                        result.Errors.Add("Thieu dieu kien upper bound: <= 45000.");
                    }
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static bool HasCustomFilter(
            XmlNode? customFiltersNode,
            XmlNamespaceManager ns,
            string expectedOperator,
            decimal expectedValue)
        {
            if (customFiltersNode == null)
            {
                return false;
            }

            var nodes = customFiltersNode.SelectNodes("x:customFilter", ns);
            if (nodes == null)
            {
                return false;
            }

            foreach (XmlNode node in nodes)
            {
                var op = node.Attributes?["operator"]?.Value ?? string.Empty;
                var valueText = node.Attributes?["val"]?.Value ?? string.Empty;
                if (!decimal.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
                {
                    continue;
                }

                if (string.Equals(op, expectedOperator, StringComparison.OrdinalIgnoreCase)
                    && value == expectedValue)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
