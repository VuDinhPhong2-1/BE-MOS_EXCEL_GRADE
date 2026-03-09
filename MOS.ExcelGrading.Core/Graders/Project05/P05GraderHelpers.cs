using OfficeOpenXml;
using System.Xml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    internal static class P05GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorksheet anySheet, string name)
        {
            return anySheet.Workbook.Worksheets.FirstOrDefault(w =>
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

        public static string NormalizeText(string? value)
        {
            return (value ?? string.Empty)
                .Trim()
                .Replace("  ", " ", StringComparison.Ordinal);
        }

        public static bool IsDifferenceFormula(string? formula, int row)
        {
            var normalized = NormalizeFormula(formula);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            if (normalized.StartsWith("(") && normalized.EndsWith(")"))
            {
                normalized = normalized[1..^1];
            }

            return string.Equals(normalized, $"F{row}-G{row}", StringComparison.Ordinal);
        }

        public static bool IsDifferenceFormulaR1C1(string? formulaR1C1)
        {
            var normalized = NormalizeFormula(formulaR1C1);
            return string.Equals(normalized, "RC[-2]-RC[-1]", StringComparison.Ordinal);
        }

        public static XmlNodeList? GetWorkbookViewNodes(ExcelWorksheet anySheet)
        {
            var workbookXml = anySheet.Workbook.WorkbookXml;
            if (workbookXml == null)
            {
                return null;
            }

            var ns = new XmlNamespaceManager(workbookXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return workbookXml.SelectNodes("//x:bookViews/x:workbookView", ns);
        }

        public static bool TryParseDoubleAttribute(XmlNode? node, string attributeName, out double value)
        {
            value = 0;
            if (node == null)
            {
                return false;
            }

            var raw = node.Attributes?[attributeName]?.Value;
            return double.TryParse(raw, out value);
        }
    }
}
