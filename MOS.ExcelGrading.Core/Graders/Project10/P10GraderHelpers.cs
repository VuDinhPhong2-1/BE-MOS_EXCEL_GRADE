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
