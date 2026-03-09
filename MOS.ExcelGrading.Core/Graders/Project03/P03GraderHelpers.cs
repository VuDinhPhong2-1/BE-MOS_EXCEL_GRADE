using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    internal static class P03GraderHelpers
    {
        public static ExcelWorksheet? GetIngredientsSheet(ExcelWorksheet studentSheet)
        {
            return studentSheet.Workbook.Worksheets.FirstOrDefault(w =>
                string.Equals((w.Name ?? string.Empty).Trim(), "Ingredients", StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeRef(string? reference)
        {
            return (reference ?? string.Empty)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Trim()
                .ToUpperInvariant();
        }

        public static string ExtractRightHeaderText(string? rightHeaderRaw)
        {
            var raw = rightHeaderRaw ?? string.Empty;
            return raw.Replace("&R", string.Empty, StringComparison.OrdinalIgnoreCase);
        }
    }
}
