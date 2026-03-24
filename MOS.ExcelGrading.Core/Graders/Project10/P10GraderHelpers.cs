using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    internal static class P10GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeRange(string? range)
        {
            var text = (range ?? string.Empty)
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
                .Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase)
                .ToUpperInvariant();
        }

        public static string NormalizeText(string? value)
        {
            return (value ?? string.Empty).Trim();
        }

        public static ExcelTable? FindTableByAddress(ExcelWorksheet sheet, string expectedAddress)
        {
            return sheet.Tables.FirstOrDefault(table =>
                IsRangeMatch(table.Address.Address, expectedAddress));
        }

        public static string JoinTableAddresses(ExcelWorksheet sheet)
        {
            return string.Join(", ", sheet.Tables.Select(t => t.Address.Address));
        }

        public static string GetChartBounds(ExcelChart chart)
        {
            var start = CellAddress(chart.From.Row + 1, chart.From.Column + 1);
            var end = CellAddress(chart.To.Row + 1, chart.To.Column + 1);
            return $"{start}:{end}";
        }

        public static bool IsChartBoundsMatch(ExcelChart chart, string expectedBounds)
        {
            return string.Equals(
                NormalizeRange(GetChartBounds(chart)),
                NormalizeRange(expectedBounds),
                StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsSeriesRangeMatch(ExcelChart chart, string expectedXRange, string expectedYRange)
        {
            var series = chart.Series.FirstOrDefault();
            if (series == null)
            {
                return false;
            }

            return IsRangeMatch(series.XSeries?.ToString(), expectedXRange)
                   && IsRangeMatch(series.Series?.ToString(), expectedYRange);
        }

        public static bool TryGetDefinedName(ExcelWorkbook workbook, string definedName, out string value)
        {
            value = string.Empty;
            var workbookXml = workbook.WorkbookXml;
            if (workbookXml == null)
            {
                return false;
            }

            var ns = new XmlNamespaceManager(workbookXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var node = workbookXml.SelectSingleNode(
                $"//x:definedNames/x:definedName[@name='{definedName}']",
                ns);

            if (node == null)
            {
                return false;
            }

            value = node.InnerText ?? string.Empty;
            return true;
        }

        public static (string ChoiceStyle, string FallbackStyle) GetChartAlternateStyles(ExcelChart chart)
        {
            var xml = chart.ChartXml;
            var choiceStyle = xml.SelectSingleNode(
                                  "/*[local-name()='chartSpace']/*[local-name()='AlternateContent']/*[local-name()='Choice']/*[local-name()='style']")
                              ?.Attributes?["val"]?.Value
                              ?? string.Empty;

            var fallbackStyle = xml.SelectSingleNode(
                                    "/*[local-name()='chartSpace']/*[local-name()='AlternateContent']/*[local-name()='Fallback']/*[local-name()='style']")
                                ?.Attributes?["val"]?.Value
                                ?? string.Empty;

            return (choiceStyle.Trim(), fallbackStyle.Trim());
        }

        public static (string ColorStyleId, string ChartStyleId) GetStyleManagerIds(ExcelChart chart)
        {
            var colorStyleId = chart.StyleManager?.ColorsXml?.DocumentElement?.Attributes?["id"]?.Value?.Trim() ?? string.Empty;
            var chartStyleId = chart.StyleManager?.StyleXml?.DocumentElement?.Attributes?["id"]?.Value?.Trim() ?? string.Empty;
            return (colorStyleId, chartStyleId);
        }

        public static int FindRowContainsText(ExcelWorksheet worksheet, int column, int startRow, int endRow, string expectedText)
        {
            for (var row = startRow; row <= endRow; row++)
            {
                var text = NormalizeText(worksheet.Cells[row, column].Text);
                if (string.Equals(text, NormalizeText(expectedText), StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            return -1;
        }

        public static string CellAddress(int row, int col)
        {
            var column = string.Empty;
            var current = col;
            while (current > 0)
            {
                current--;
                column = (char)('A' + (current % 26)) + column;
                current /= 26;
            }

            return $"{column}{row}";
        }
    }
}
