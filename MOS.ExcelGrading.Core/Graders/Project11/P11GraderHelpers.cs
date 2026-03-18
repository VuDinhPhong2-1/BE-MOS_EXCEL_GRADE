using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    internal static class P11GraderHelpers
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

        public static bool IsRangeMatch(string? actual, string expected)
        {
            return string.Equals(
                NormalizeRange(actual),
                NormalizeRange(expected),
                StringComparison.OrdinalIgnoreCase);
        }

        public static XmlNamespaceManager CreateWorkbookNamespaceManager(XmlDocument workbookXml)
        {
            var ns = new XmlNamespaceManager(workbookXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return ns;
        }

        public static int GetSheetIndex0Based(ExcelWorkbook workbook, string sheetName)
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (string.Equals(workbook.Worksheets[i].Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
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

        public static string NormalizeUrl(string? url)
        {
            return (url ?? string.Empty)
                .Trim()
                .TrimEnd('/')
                .ToLowerInvariant();
        }
    }
}
