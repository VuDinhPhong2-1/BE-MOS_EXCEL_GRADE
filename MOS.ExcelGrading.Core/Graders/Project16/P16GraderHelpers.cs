using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project16
{
    internal static class P16GraderHelpers
    {
        private static readonly Regex SqrefPartRegex = new(
            @"^(?<c1>[A-Z]+)(?<r1>\d+)(:(?<c2>[A-Z]+)(?<r2>\d+))?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
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
            return string.Equals(NormalizeRange(actual), NormalizeRange(expected), StringComparison.OrdinalIgnoreCase);
        }

        public static bool SqrefContainsRange(string? sqref, string expectedRange)
        {
            var normalizedExpected = NormalizeRange(expectedRange);
            var parts = (sqref ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return parts.Any(part =>
                string.Equals(NormalizeRange(part), normalizedExpected, StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeIdentifier(string? value)
        {
            var text = (value ?? string.Empty).Trim().ToUpperInvariant();
            return new string(text.Where(char.IsLetterOrDigit).ToArray());
        }

        public static bool IsQuantityColumnName(string? value)
        {
            var normalized = NormalizeIdentifier(value);
            return normalized == "QUANTITY" || normalized == "QTY";
        }

        public static bool SqrefTargetsColumn(string? sqref, int expectedColumn, int dataStartRow, int dataEndRow)
        {
            var expectedColumnText = new string(ExcelCellBase.GetAddress(1, expectedColumn).Where(char.IsLetter).ToArray()).ToUpperInvariant();
            var parts = (sqref ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            foreach (var part in parts)
            {
                var normalizedPart = NormalizeRange(part);
                var match = SqrefPartRegex.Match(normalizedPart);
                if (!match.Success)
                {
                    continue;
                }

                var col1 = match.Groups["c1"].Value.ToUpperInvariant();
                var col2 = (match.Groups["c2"].Success ? match.Groups["c2"].Value : match.Groups["c1"].Value).ToUpperInvariant();
                if (!string.Equals(col1, expectedColumnText, StringComparison.Ordinal)
                    || !string.Equals(col2, expectedColumnText, StringComparison.Ordinal))
                {
                    continue;
                }

                var row1 = int.Parse(match.Groups["r1"].Value);
                var row2 = int.Parse((match.Groups["r2"].Success ? match.Groups["r2"].Value : match.Groups["r1"].Value));
                var minRow = Math.Min(row1, row2);
                var maxRow = Math.Max(row1, row2);

                if (maxRow >= dataStartRow && minRow <= dataEndRow)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool TryFindColumnByHeader(
            ExcelWorksheet worksheet,
            Func<string, bool> headerMatcher,
            out int headerRow,
            out int columnIndex)
        {
            headerRow = -1;
            columnIndex = -1;

            var maxHeaderRow = Math.Min(10, worksheet.Dimension.End.Row);
            for (var row = 1; row <= maxHeaderRow; row++)
            {
                for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var text = (worksheet.Cells[row, col].Text ?? string.Empty).Trim();
                    if (headerMatcher(text))
                    {
                        headerRow = row;
                        columnIndex = col;
                        return true;
                    }
                }
            }

            return false;
        }

        public static int GetLastDataRowInColumn(ExcelWorksheet worksheet, int columnIndex, int startRow)
        {
            var lastRow = startRow;
            for (var row = startRow; row <= worksheet.Dimension.End.Row; row++)
            {
                var cell = worksheet.Cells[row, columnIndex];
                if (!string.IsNullOrWhiteSpace(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula))
                {
                    lastRow = row;
                }
            }

            return lastRow;
        }

        public static (string ColorStyleId, string ChartStyleId) GetDrawingStyleIds(object drawing)
        {
            var styleManager = drawing.GetType().GetProperty("StyleManager")?.GetValue(drawing);
            if (styleManager == null)
            {
                return (string.Empty, string.Empty);
            }

            var colorsXml = styleManager.GetType().GetProperty("ColorsXml")?.GetValue(styleManager) as XmlDocument;
            var styleXml = styleManager.GetType().GetProperty("StyleXml")?.GetValue(styleManager) as XmlDocument;
            var colorStyleId = colorsXml?.DocumentElement?.Attributes?["id"]?.Value?.Trim() ?? string.Empty;
            var chartStyleId = styleXml?.DocumentElement?.Attributes?["id"]?.Value?.Trim() ?? string.Empty;
            return (colorStyleId, chartStyleId);
        }
    }
}

