using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    internal static class P04GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorksheet studentSheet, string name)
        {
            return studentSheet.Workbook.Worksheets.FirstOrDefault(w =>
                string.Equals((w.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeAddress(string? address)
        {
            return (address ?? string.Empty)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Trim()
                .ToUpperInvariant();
        }

        public static bool IsWidth12(double width)
        {
            // Enforce exact column width corresponding to Excel UI "12".
            // In workbook XML this is typically persisted as 12.7109375.
            return Math.Abs(width - 12.7109375d) <= 0.01d;
        }

        public static bool TryParseSection(string? text, out int section)
        {
            section = 0;
            var raw = (text ?? string.Empty).Trim();
            if (int.TryParse(raw, out var direct))
            {
                section = direct;
                return true;
            }

            var digits = new string(raw.Where(char.IsDigit).ToArray());
            if (digits.Length == 0)
            {
                return false;
            }

            if (int.TryParse(digits, out var parsed))
            {
                section = parsed;
                return true;
            }

            return false;
        }

        public static bool IsGraduationChartSheetName(string? name)
        {
            var n = (name ?? string.Empty).Trim();
            return string.Equals(n, "Graduation Chart", StringComparison.Ordinal);
        }
    }
}
