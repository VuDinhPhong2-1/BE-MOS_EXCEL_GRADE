using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project09
{
    public class P09T4Grader : ITaskGrader
    {
        public string TaskId => "P09-T4";
        public string TaskName => "Loc cot Total: tu 34,000 den 45,000";
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
                var ws = P09GraderHelpers.GetSheet(studentSheet.Workbook, "Data");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Data'.");
                    return result;
                }

                decimal score = 0m;

                var filterAddress = ws.AutoFilter?.Address?.Address ?? string.Empty;
                var normalizedFilterAddress = P09GraderHelpers.NormalizeRange(filterAddress);
                if (string.Equals(normalizedFilterAddress, "A17:K29", StringComparison.OrdinalIgnoreCase)
    || string.Equals(normalizedFilterAddress, "K17:K29", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add("AutoFilter dung vung A17:K29.");
                }
                else
                {
                    result.Errors.Add($"AutoFilter chua dúng. Hi?n t?i: '{filterAddress}'.");
                }

                var ns = P09GraderHelpers.CreateWorksheetNamespaceManager(ws.WorksheetXml);
                var autoFilterNode = ws.WorksheetXml.SelectSingleNode("//x:autoFilter", ns);
                if (autoFilterNode == null)
                {
                    result.Errors.Add("Không tìm th?y node autoFilter trong XML.");
                    result.Score = score;
                    return result;
                }

                var autoFilterAddress = ws.AutoFilter?.Address;
                if (autoFilterAddress == null)
                {
                    result.Errors.Add("Không d?c du?c dia chi AutoFilter de xac dinh c?t Total.");
                    result.Score = score;
                    return result;
                }

                var totalColumnIds = GetTotalColumnIds(ws, autoFilterAddress);
                if (totalColumnIds.Count == 0)
                {
                    result.Errors.Add("Không tìm th?y c?t 'Total' trong hàng tieu de AutoFilter.");
                    result.Score = score;
                    return result;
                }

                var filterColumnNodes = autoFilterNode.SelectNodes("x:filterColumn", ns);
                if (filterColumnNodes == null || filterColumnNodes.Count == 0)
                {
                    result.Errors.Add("Không tìm th?y filterColumn trong autoFilter.");
                    result.Score = score;
                    return result;
                }

                var totalFilterColumns = new List<XmlNode>();
                foreach (XmlNode filterColumn in filterColumnNodes)
                {
                    var colIdText = filterColumn.Attributes?["colId"]?.Value ?? string.Empty;
                    if (!int.TryParse(colIdText, out var colId))
                    {
                        continue;
                    }

                    if (totalColumnIds.Contains(colId))
                    {
                        totalFilterColumns.Add(filterColumn);
                    }
                }

                if (totalFilterColumns.Count > 0)
                {
                    score += 1m;
                    result.Details.Add("Filter dat dung tren it nhat mot c?t Total.");
                }
                else
                {
                    result.Errors.Add("Không tìm th?y filterColumn tren cac c?t Total.");
                }

                var hasMatchedConditions = false;
                var hasAndInAnyTotalColumn = false;
                var hasLowerBoundInAnyTotalColumn = false;
                var hasUpperBoundInAnyTotalColumn = false;

                foreach (var totalFilterColumn in totalFilterColumns)
                {
                    var customFiltersNode = totalFilterColumn.SelectSingleNode("x:customFilters", ns);
                    var andAttr = customFiltersNode?.Attributes?["and"]?.Value ?? string.Empty;
                    var hasAnd = string.Equals(andAttr, "1", StringComparison.OrdinalIgnoreCase)
                                 || string.Equals(andAttr, "true", StringComparison.OrdinalIgnoreCase);
                    var hasLowerBound = HasCustomFilter(customFiltersNode, ns, "greaterThanOrEqual", 34000m);
                    var hasUpperBound = HasCustomFilter(customFiltersNode, ns, "lessThanOrEqual", 45000m);

                    hasAndInAnyTotalColumn = hasAndInAnyTotalColumn || hasAnd;
                    hasLowerBoundInAnyTotalColumn = hasLowerBoundInAnyTotalColumn || hasLowerBound;
                    hasUpperBoundInAnyTotalColumn = hasUpperBoundInAnyTotalColumn || hasUpperBound;

                    if (hasAnd && hasLowerBound && hasUpperBound)
                    {
                        hasMatchedConditions = true;
                        break;
                    }
                }

                if (hasMatchedConditions)
                {
                    score += 2m;
                    result.Details.Add("Dieu kien filter dung: >= 34000 va <= 45000.");
                }
                else
                {
                    if (!hasAndInAnyTotalColumn)
                    {
                        result.Errors.Add("customFilters chua dat di?u ki?n AND.");
                    }

                    if (!hasLowerBoundInAnyTotalColumn)
                    {
                        result.Errors.Add("Thieu di?u ki?n lower bound: >= 34000.");
                    }

                    if (!hasUpperBoundInAnyTotalColumn)
                    {
                        result.Errors.Add("Thieu di?u ki?n upper bound: <= 45000.");
                    }

                    if (hasAndInAnyTotalColumn && hasLowerBoundInAnyTotalColumn && hasUpperBoundInAnyTotalColumn)
                    {
                        result.Errors.Add("Cac di?u ki?n filter chua nam cung mot c?t Total.");
                    }
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
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

        private static List<int> GetTotalColumnIds(ExcelWorksheet ws, ExcelAddressBase autoFilterAddress)
        {
            var result = new List<int>();
            var headerRow = autoFilterAddress.Start.Row;
            var startColumn = autoFilterAddress.Start.Column;
            var endColumn = autoFilterAddress.End.Column;

            for (var column = startColumn; column <= endColumn; column++)
            {
                var headerText = ws.Cells[headerRow, column].Text?.Trim() ?? string.Empty;
                if (!string.Equals(headerText, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                result.Add(column - startColumn);
            }

            return result;
        }
    }
}

// minor-sync: non-functional graders update




