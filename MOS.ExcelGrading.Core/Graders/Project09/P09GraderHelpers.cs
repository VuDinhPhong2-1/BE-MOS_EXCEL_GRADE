using System.Xml;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    internal static class P09GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static XmlNamespaceManager CreateWorksheetNamespaceManager(XmlDocument xml)
        {
            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return ns;
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
