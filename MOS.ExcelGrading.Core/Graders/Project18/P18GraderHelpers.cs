using System.Globalization;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project18
{
    internal static class P18GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string sheetName)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), sheetName.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static ExcelNamedRange? FindNamedRange(ExcelWorkbook workbook, string rangeName)
        {
            var workbookLevelName = workbook.Names.FirstOrDefault(name =>
                string.Equals(name.Name, rangeName, StringComparison.OrdinalIgnoreCase));
            if (workbookLevelName != null)
            {
                return workbookLevelName;
            }

            foreach (var worksheet in workbook.Worksheets)
            {
                var worksheetLevelName = worksheet.Names.FirstOrDefault(name =>
                    string.Equals(name.Name, rangeName, StringComparison.OrdinalIgnoreCase));
                if (worksheetLevelName != null)
                {
                    return worksheetLevelName;
                }
            }

            return null;
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

        public static string NormalizeIdentifier(string? value)
        {
            var text = (value ?? string.Empty).Trim().ToUpperInvariant();
            return new string(text.Where(char.IsLetterOrDigit).ToArray());
        }

        public static bool IsTwoDecimalNumberFormat(ExcelRangeBase cell)
        {
            var numberFormatId = cell.Style.Numberformat.NumFmtID;
            if (numberFormatId == 2)
            {
                return true;
            }

            var format = (cell.Style.Numberformat.Format ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(format))
            {
                return false;
            }

            var normalizedFormat = format.Replace(" ", string.Empty, StringComparison.Ordinal);
            var positiveSection = normalizedFormat.Split(';')[0];
            if (string.IsNullOrWhiteSpace(positiveSection))
            {
                return false;
            }

            if (!positiveSection.Contains("0.00", StringComparison.Ordinal))
            {
                return false;
            }

            return !positiveSection.Contains("0.000", StringComparison.Ordinal);
        }

        public static bool CellIsEmpty(ExcelRangeBase cell)
        {
            return string.IsNullOrWhiteSpace(cell.Text)
                   && string.IsNullOrWhiteSpace(cell.Formula)
                   && cell.Value == null;
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
                    return decimal.TryParse(
                        cell.Value.ToString(),
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out value)
                           || decimal.TryParse(
                               cell.Value.ToString(),
                               NumberStyles.Any,
                               CultureInfo.CurrentCulture,
                               out value);
            }
        }

        public static ExcelTable? FindTableByHeaders(ExcelWorksheet worksheet, params string[] requiredHeaders)
        {
            var normalizedRequiredHeaders = requiredHeaders
                .Select(NormalizeIdentifier)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var table in worksheet.Tables)
            {
                var normalizedTableHeaders = table.Columns
                    .Select(column => NormalizeIdentifier(column.Name))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                if (normalizedRequiredHeaders.All(normalizedTableHeaders.Contains))
                {
                    return table;
                }
            }

            return null;
        }

        public static bool IsAverageFormulaForMonthlyRange(
            string? formula,
            int row,
            int januaryColumn,
            int aprilColumn)
        {
            var normalized = NormalizeFormula(formula);
            if (!normalized.Contains("AVERAGE(", StringComparison.Ordinal))
            {
                return false;
            }

            var explicitRange = NormalizeRange(
                $"{ExcelCellBase.GetAddress(row, januaryColumn)}:{ExcelCellBase.GetAddress(row, aprilColumn)}");
            if (normalized.Contains(explicitRange, StringComparison.Ordinal))
            {
                return true;
            }

            if (normalized.Contains("[JANUARY]:[APRIL]", StringComparison.Ordinal))
            {
                return true;
            }

            var allMonthCellsReferenced = true;
            for (var col = januaryColumn; col <= aprilColumn; col++)
            {
                var monthAddress = ExcelCellBase.GetAddress(row, col).ToUpperInvariant();
                if (!normalized.Contains(monthAddress, StringComparison.Ordinal))
                {
                    allMonthCellsReferenced = false;
                    break;
                }
            }

            return allMonthCellsReferenced;
        }

        public static bool HasExactQuotedLiteral(string? formula, string expectedLiteral)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return false;
            }

            var matches = Regex.Matches(formula, "\"([^\"]*)\"");
            foreach (Match match in matches)
            {
                var literal = match.Groups[1].Value;
                if (string.Equals(literal, expectedLiteral, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool HasSpaceBeforeDomainLiteral(string? formula, string domainLiteral)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return false;
            }

            var escapedDomain = Regex.Escape(domainLiteral.TrimStart('@'));
            return Regex.IsMatch(
                formula,
                "\"\\s+@" + escapedDomain + "\"",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        public static List<ChartSeriesInfo> ReadChartSeriesInfo(ExcelChart chart)
        {
            var infos = new List<ChartSeriesInfo>();

            foreach (var series in chart.Series)
            {
                infos.Add(new ChartSeriesInfo
                {
                    HeaderAddress = NormalizeRange(series.HeaderAddress?.Address),
                    CategoryAddress = NormalizeRange(series.XSeries),
                    ValueAddress = NormalizeRange(series.Series)
                });
            }

            return infos;
        }

        internal sealed class ChartSeriesInfo
        {
            public string HeaderAddress { get; init; } = string.Empty;
            public string CategoryAddress { get; init; } = string.Empty;
            public string ValueAddress { get; init; } = string.Empty;
        }
    }
}

