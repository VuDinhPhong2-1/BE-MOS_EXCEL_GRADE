using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project14
{
    internal static class P14GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeFormula(string? formula)
        {
            var normalized = (formula ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Replace(";", ",", StringComparison.Ordinal)
                .Replace("@", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();

            // Excel may store table formulas in equivalent forms:
            // Table1[[#This Row],[Col]] or Table1[Col].
            return normalized
                .Replace("[[#THISROW],[", "[", StringComparison.Ordinal)
                .Replace("]]", "]", StringComparison.Ordinal);
        }

        public static string NormalizePrintArea(string? value)
        {
            return (value ?? string.Empty)
                .Trim()
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();
        }

        public static string GetSingleFilterValue(ExcelTable table, int colId)
        {
            var ns = new XmlNamespaceManager(table.TableXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var node = table.TableXml.SelectSingleNode(
                $"//x:autoFilter/x:filterColumn[@colId='{colId}']/x:filters/x:filter",
                ns);
            return node?.Attributes?["val"]?.Value?.Trim() ?? string.Empty;
        }

        public static ExcelTable? FindTableByAddress(ExcelWorksheet worksheet, string expectedAddress)
        {
            return worksheet.Tables.FirstOrDefault(table =>
                string.Equals(
                    NormalizePrintArea(table.Address.Address),
                    NormalizePrintArea(expectedAddress),
                    StringComparison.OrdinalIgnoreCase));
        }
    }
}

