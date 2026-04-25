using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project12
{
    internal static class P12GraderHelpers
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
            return string.Equals(NormalizeRange(actual), NormalizeRange(expected), StringComparison.OrdinalIgnoreCase);
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

        public static bool HasMergeRange(ExcelWorksheet worksheet, string expectedRange)
        {
            return worksheet.MergedCells.Any(range => IsRangeMatch(range, expectedRange));
        }

        public static ExcelTable? FindTableByAddress(ExcelWorksheet worksheet, string expectedAddress)
        {
            return worksheet.Tables.FirstOrDefault(table => IsRangeMatch(table.Address.Address, expectedAddress));
        }

        public static string GetSingleFilterValue(ExcelTable table, int colId)
        {
            var tableXml = table.TableXml;
            var ns = new XmlNamespaceManager(tableXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var filterNode = tableXml.SelectSingleNode(
                $"//x:autoFilter/x:filterColumn[@colId='{colId}']/x:filters/x:filter",
                ns);
            return filterNode?.Attributes?["val"]?.Value?.Trim() ?? string.Empty;
        }
    }
}

