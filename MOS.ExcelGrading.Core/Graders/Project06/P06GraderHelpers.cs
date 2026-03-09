using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project06
{
    internal static class P06GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorksheet studentSheet, string name)
        {
            return studentSheet.Workbook.Worksheets.FirstOrDefault(w =>
                string.Equals((w.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeFormula(string? formula)
        {
            return (formula ?? string.Empty)
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase)
                .ToUpperInvariant()
                .Trim();
        }

        public static string NormalizeAddress(string? address)
        {
            var text = (address ?? string.Empty)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Trim();
            var excl = text.LastIndexOf('!');
            if (excl >= 0 && excl + 1 < text.Length)
            {
                text = text[(excl + 1)..];
            }
            return text.Trim('\'').ToUpperInvariant();
        }

        public static decimal ToDecimal(object? value, string? text = null)
        {
            if (value is decimal d) return d;
            if (value is double db) return Convert.ToDecimal(db);
            if (value is float f) return Convert.ToDecimal(f);
            if (value is int i) return i;
            if (value is long l) return l;

            if (decimal.TryParse(text ?? string.Empty, out var parsed))
            {
                return parsed;
            }

            var cleaned = (text ?? string.Empty)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(",", string.Empty, StringComparison.Ordinal)
                .Trim();

            return decimal.TryParse(cleaned, out var cleanedParsed) ? cleanedParsed : 0m;
        }
    }
}
