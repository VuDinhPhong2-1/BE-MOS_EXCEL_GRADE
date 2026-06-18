using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project20
{
    internal static class P20GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeRange(string? value)
        {
            var text = (value ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal);

            var excl = text.LastIndexOf('!');
            if (excl >= 0 && excl + 1 < text.Length)
            {
                text = text[(excl + 1)..];
            }

            return text.ToUpperInvariant();
        }

        public static bool IsRangeMatch(string? actual, string expected)
        {
            return string.Equals(
                NormalizeRange(actual),
                NormalizeRange(expected),
                StringComparison.OrdinalIgnoreCase);
        }

        public static string NormalizeFormula(string? formula)
        {
            return (formula ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace(";", ",", StringComparison.Ordinal)
                .Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Replace("@", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();
        }

        public static string NormalizeText(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Replace('\u00A0', ' ').Trim();
            return string.Join(" ", normalized.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        }

        public static string NormalizeIdentifier(string? value)
        {
            var text = NormalizeText(value).ToUpperInvariant();
            return new string(text.Where(char.IsLetterOrDigit).ToArray());
        }

        public static bool TryGetNumericValue(ExcelRangeBase cell, out decimal value)
        {
            value = 0m;
            if (cell.Value == null)
            {
                return false;
            }

            switch (cell.Value)
            {
                case decimal decimalValue:
                    value = decimalValue;
                    return true;
                case double doubleValue:
                    value = Convert.ToDecimal(doubleValue, CultureInfo.InvariantCulture);
                    return true;
                case float floatValue:
                    value = Convert.ToDecimal(floatValue, CultureInfo.InvariantCulture);
                    return true;
                case int intValue:
                    value = intValue;
                    return true;
                case long longValue:
                    value = longValue;
                    return true;
                case string stringValue:
                    return decimal.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out value)
                           || decimal.TryParse(stringValue, NumberStyles.Any, CultureInfo.CurrentCulture, out value);
                default:
                    return decimal.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out value)
                           || decimal.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out value);
            }
        }

        public static ExcelTable? FindTable(
            ExcelWorksheet worksheet,
            string? preferredDisplayName,
            params string[] requiredHeaders)
        {
            if (!string.IsNullOrWhiteSpace(preferredDisplayName))
            {
                var byName = worksheet.Tables.FirstOrDefault(table =>
                    string.Equals(table.Name?.Trim(), preferredDisplayName.Trim(), StringComparison.OrdinalIgnoreCase));
                if (byName != null)
                {
                    return byName;
                }
            }

            var required = requiredHeaders
                .Select(NormalizeIdentifier)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var table in worksheet.Tables)
            {
                var headers = table.Columns
                    .Select(column => NormalizeIdentifier(column.Name))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                if (required.All(headers.Contains))
                {
                    return table;
                }
            }

            return null;
        }

        public static bool TryGetColumnIndex(ExcelTable table, string headerName, out int columnIndex)
        {
            columnIndex = -1;
            var normalizedHeader = NormalizeIdentifier(headerName);

            for (var i = 0; i < table.Columns.Count; i++)
            {
                if (string.Equals(
                    NormalizeIdentifier(table.Columns[i].Name),
                    normalizedHeader,
                    StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = table.Address.Start.Column + i;
                    return true;
                }
            }

            return false;
        }

        public static bool TryGetTableSortState(ExcelTable table, out string sortRef, out List<string> sortConditionRefs)
        {
            sortRef = string.Empty;
            sortConditionRefs = new List<string>();

            var xml = table.TableXml;
            if (xml == null)
            {
                return false;
            }

            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var sortNode = xml.SelectSingleNode("/x:table/x:sortState", ns);
            if (sortNode == null)
            {
                return false;
            }

            sortRef = sortNode.Attributes?["ref"]?.Value ?? string.Empty;
            var conditions = sortNode.SelectNodes("x:sortCondition", ns);
            if (conditions != null)
            {
                foreach (XmlNode condition in conditions)
                {
                    var conditionRef = condition.Attributes?["ref"]?.Value ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(conditionRef))
                    {
                        sortConditionRefs.Add(conditionRef);
                    }
                }
            }

            return true;
        }

        public static int CompareSortText(string left, string right)
        {
            return CultureInfo.InvariantCulture.CompareInfo.Compare(
                NormalizeText(left),
                NormalizeText(right),
                CompareOptions.IgnoreCase | CompareOptions.IgnoreNonSpace | CompareOptions.IgnoreWidth);
        }

        public static string GetChartBounds(ExcelChart chart)
        {
            var start = ExcelCellBase.GetAddress(chart.From.Row + 1, chart.From.Column + 1);
            var end = ExcelCellBase.GetAddress(chart.To.Row + 1, chart.To.Column + 1);
            return $"{start}:{end}";
        }

        public static bool IsClusteredColumnChart(ExcelChart chart)
        {
            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            var barChart = xml.SelectSingleNode("//c:chart/c:plotArea/c:barChart", ns);
            if (barChart == null)
            {
                return false;
            }

            var barDir = barChart.SelectSingleNode("c:barDir", ns)?.Attributes?["val"]?.Value ?? string.Empty;
            var grouping = barChart.SelectSingleNode("c:grouping", ns)?.Attributes?["val"]?.Value ?? string.Empty;

            return string.Equals(barDir, "col", StringComparison.OrdinalIgnoreCase)
                   && string.Equals(grouping, "clustered", StringComparison.OrdinalIgnoreCase);
        }

        public static bool TryGetFirstSeriesRanges(ExcelChart chart, out string headerRange, out string categoryRange, out string valueRange)
        {
            headerRange = string.Empty;
            categoryRange = string.Empty;
            valueRange = string.Empty;

            var series = chart.Series.FirstOrDefault();
            if (series == null)
            {
                return false;
            }

            headerRange = NormalizeRange(series.HeaderAddress?.Address);
            categoryRange = NormalizeRange(series.XSeries?.ToString());
            valueRange = NormalizeRange(series.Series?.ToString());
            return true;
        }

        public static bool HasDataTable(ExcelChart chart, out bool showKeys)
        {
            showKeys = false;

            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            var dTable = xml.SelectSingleNode("//c:chart/c:plotArea/c:dTable", ns);
            if (dTable == null)
            {
                return false;
            }

            var showKeysValue = dTable.SelectSingleNode("c:showKeys", ns)?.Attributes?["val"]?.Value
                                ?? dTable.SelectSingleNode("c:showLegendKeys", ns)?.Attributes?["val"]?.Value
                                ?? "0";

            showKeys = !string.Equals(showKeysValue, "0", StringComparison.OrdinalIgnoreCase);
            return true;
        }

        public static bool IsLegendHidden(ExcelChart chart)
        {
            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            var legend = xml.SelectSingleNode("//c:chart/c:legend", ns);
            if (legend == null)
            {
                return true;
            }

            var deleteValue = legend.SelectSingleNode("c:delete", ns)?.Attributes?["val"]?.Value ?? string.Empty;
            return string.Equals(deleteValue, "1", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsChartTitleBlank(ExcelChart chart)
        {
            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var titleNode = xml.SelectSingleNode("//c:chart/c:title", ns);
            if (titleNode == null)
            {
                return true;
            }

            var textNodes = titleNode.SelectNodes(".//a:t", ns);
            if (textNodes == null || textNodes.Count == 0)
            {
                return true;
            }

            var text = string.Join(
                string.Empty,
                textNodes.Cast<XmlNode>().Select(node => node.InnerText ?? string.Empty));
            return string.IsNullOrWhiteSpace(text);
        }

        public static bool FormulaLooksLikeMaxOverAirMiles(string? formula, int airMilesColumn, int startRow, int endRow)
        {
            var normalized = NormalizeFormula(formula);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            if (!normalized.StartsWith("MAX(", StringComparison.Ordinal) || !normalized.EndsWith(")", StringComparison.Ordinal))
            {
                return false;
            }

            var argument = normalized[4..^1];
            if (argument.Contains(",", StringComparison.Ordinal))
            {
                return false;
            }

            if (argument.Contains("D:D", StringComparison.Ordinal)
                || Regex.IsMatch(argument, @"[A-Z]+:[A-Z]+", RegexOptions.CultureInvariant))
            {
                return false;
            }

            var expectedRange = $"{ExcelCellBase.GetAddress(startRow, airMilesColumn)}:{ExcelCellBase.GetAddress(endRow, airMilesColumn)}";
            var normalizedExpected = NormalizeRange(expectedRange);

            return argument.Contains("TABLE14[AIRMILES]", StringComparison.Ordinal)
                   || argument.Contains("[AIRMILES]", StringComparison.Ordinal)
                   || string.Equals(NormalizeRange(argument), normalizedExpected, StringComparison.OrdinalIgnoreCase);
        }
    }
}


