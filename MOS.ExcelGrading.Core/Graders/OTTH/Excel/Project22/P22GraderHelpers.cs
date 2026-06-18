using System.Globalization;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project22
{
    internal static class P22GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string sheetName)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), sheetName.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static ExcelChartsheet? GetChartSheet(ExcelWorkbook workbook, string sheetName)
        {
            return workbook.Worksheets
                .FirstOrDefault(sheet =>
                    sheet is ExcelChartsheet
                    && string.Equals((sheet.Name ?? string.Empty).Trim(), sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                as ExcelChartsheet;
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

        public static string NormalizeRange(string? rangeAddress)
        {
            var normalized = (rangeAddress ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal);

            var exclamationIndex = normalized.LastIndexOf('!');
            if (exclamationIndex >= 0 && exclamationIndex + 1 < normalized.Length)
            {
                normalized = normalized[(exclamationIndex + 1)..];
            }

            return normalized.ToUpperInvariant();
        }

        public static bool IsRangeMatch(string? actualRange, string expectedRange)
        {
            return string.Equals(
                NormalizeRange(actualRange),
                NormalizeRange(expectedRange),
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
            string? preferredName,
            params string[] requiredHeaders)
        {
            if (!string.IsNullOrWhiteSpace(preferredName))
            {
                var byName = worksheet.Tables.FirstOrDefault(table =>
                    string.Equals(table.Name?.Trim(), preferredName.Trim(), StringComparison.OrdinalIgnoreCase));
                if (byName != null)
                {
                    return byName;
                }
            }

            var normalizedRequiredHeaders = requiredHeaders
                .Select(NormalizeIdentifier)
                .Where(header => !string.IsNullOrWhiteSpace(header))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var table in worksheet.Tables)
            {
                var normalizedHeaders = table.Columns
                    .Select(column => NormalizeIdentifier(column.Name))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                if (normalizedRequiredHeaders.All(normalizedHeaders.Contains))
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

        public static string GetTableDisplayName(ExcelTable table)
        {
            var documentElement = table.TableXml?.DocumentElement;
            if (documentElement == null)
            {
                return string.Empty;
            }

            return documentElement.GetAttribute("displayName") ?? string.Empty;
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

        public static bool IsLegendHidden(ExcelChart chart)
        {
            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            var legendNode = xml.SelectSingleNode("//c:chart/c:legend", ns);
            if (legendNode == null)
            {
                return true;
            }

            var deleteValue = legendNode.SelectSingleNode("c:delete", ns)?.Attributes?["val"]?.Value ?? string.Empty;
            return string.Equals(deleteValue, "1", StringComparison.OrdinalIgnoreCase);
        }

        public static bool TryGetDataLabelSettings(ExcelChart chart, out DataLabelSettings settings)
        {
            settings = new DataLabelSettings();

            var xml = chart.ChartXml;
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            var labelsNode =
                xml.SelectSingleNode("//c:chart/c:plotArea//c:ser/c:dLbls", ns)
                ?? xml.SelectSingleNode("//c:chart/c:plotArea//c:barChart/c:dLbls", ns)
                ?? xml.SelectSingleNode("//c:chart/c:plotArea//c:bar3DChart/c:dLbls", ns);

            if (labelsNode == null)
            {
                return false;
            }

            settings = new DataLabelSettings
            {
                ShowLegendKey = ReadLabelFlag(labelsNode, "showLegendKey"),
                ShowValue = ReadLabelFlag(labelsNode, "showVal"),
                ShowCategoryName = ReadLabelFlag(labelsNode, "showCatName"),
                ShowSeriesName = ReadLabelFlag(labelsNode, "showSerName"),
                ShowPercent = ReadLabelFlag(labelsNode, "showPercent"),
                ShowBubbleSize = ReadLabelFlag(labelsNode, "showBubbleSize"),
                Position = labelsNode.SelectSingleNode("c:dLblPos", ns)?.Attributes?["val"]?.Value ?? string.Empty
            };

            return true;
        }

        private static bool ReadLabelFlag(XmlNode labelsNode, string nodeName)
        {
            var node = labelsNode.ChildNodes
                .Cast<XmlNode>()
                .FirstOrDefault(child => string.Equals(child.LocalName, nodeName, StringComparison.OrdinalIgnoreCase));
            if (node == null)
            {
                return false;
            }

            var val = node.Attributes?["val"]?.Value;
            return string.IsNullOrWhiteSpace(val) || !string.Equals(val, "0", StringComparison.OrdinalIgnoreCase);
        }

        internal sealed class DataLabelSettings
        {
            public bool ShowLegendKey { get; init; }
            public bool ShowValue { get; init; }
            public bool ShowCategoryName { get; init; }
            public bool ShowSeriesName { get; init; }
            public bool ShowPercent { get; init; }
            public bool ShowBubbleSize { get; init; }
            public string Position { get; init; } = string.Empty;
        }
    }
}

