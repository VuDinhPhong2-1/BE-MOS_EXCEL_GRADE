using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project13
{
    internal static class P13GraderHelpers
    {
        public static ExcelWorksheet? GetSheet(ExcelWorkbook workbook, string name)
        {
            return workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals((sheet.Name ?? string.Empty).Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeFormula(string? formula)
        {
            return (formula ?? string.Empty)
                .Trim()
                .Replace("=", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Replace(";", ",", StringComparison.Ordinal)
                .Replace("@", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();
        }

        public static bool FormulaMatchesAny(string? actualFormula, params string[] acceptedFormulas)
        {
            var actual = NormalizeFormula(actualFormula);
            if (string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            return acceptedFormulas.Any(accepted =>
                string.Equals(actual, NormalizeFormula(accepted), StringComparison.Ordinal));
        }

        public static bool CellsHaveEquivalentValue(ExcelRangeBase studentCell, ExcelRangeBase answerCell)
        {
            var studentText = (studentCell.Text ?? string.Empty).Trim();
            var answerText = (answerCell.Text ?? string.Empty).Trim();

            if (string.Equals(studentText, answerText, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (TryGetDouble(studentCell.Value, out var studentNumeric)
                && TryGetDouble(answerCell.Value, out var answerNumeric))
            {
                return Math.Abs(studentNumeric - answerNumeric) < 0.000001d;
            }

            if (TryParseDouble(studentText, out studentNumeric)
                && TryParseDouble(answerText, out answerNumeric))
            {
                return Math.Abs(studentNumeric - answerNumeric) < 0.000001d;
            }

            return false;
        }

        public static bool CellMatchesExpected(ExcelRangeBase cell, string expectedText)
        {
            var actualText = (cell.Text ?? string.Empty).Trim();
            var normalizedExpected = (expectedText ?? string.Empty).Trim();
            if (string.Equals(actualText, normalizedExpected, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (TryGetDouble(cell.Value, out var actualNumeric)
                && TryParseDouble(normalizedExpected, out var expectedNumeric))
            {
                return Math.Abs(actualNumeric - expectedNumeric) < 0.000001d;
            }

            return TryParseDouble(actualText, out actualNumeric)
                && TryParseDouble(normalizedExpected, out expectedNumeric)
                && Math.Abs(actualNumeric - expectedNumeric) < 0.000001d;
        }

        public static HashSet<int> GetManualRowBreakIds(ExcelWorksheet worksheet)
        {
            var ids = new HashSet<int>();
            var ns = CreateWorksheetNamespaceManager(worksheet.WorksheetXml);
            var breakNodes = worksheet.WorksheetXml.SelectNodes("//x:rowBreaks/x:brk[@man='1']", ns);
            if (breakNodes == null)
            {
                return ids;
            }

            foreach (XmlNode node in breakNodes)
            {
                var idText = node.Attributes?["id"]?.Value ?? string.Empty;
                if (int.TryParse(idText, out var id))
                {
                    ids.Add(id);
                }
            }

            return ids;
        }

        public static HashSet<int> GetDetectedRowBreakIds(ExcelWorksheet worksheet)
        {
            var ids = new HashSet<int>();

            // Prefer EPPlus row metadata when available.
            var scanEndRow = Math.Max(worksheet.Dimension?.End.Row ?? 0, 300);
            for (var row = 1; row <= scanEndRow; row++)
            {
                if (worksheet.Row(row).PageBreak)
                {
                    ids.Add(row);
                }
            }

            if (ids.Count > 0)
            {
                return ids;
            }

            // XML fallback: read manual break nodes; if unavailable, read all row break nodes.
            var ns = CreateWorksheetNamespaceManager(worksheet.WorksheetXml);
            var rowBreaksNode = worksheet.WorksheetXml.SelectSingleNode("//x:rowBreaks", ns);
            var manualBreakNodes = rowBreaksNode?.SelectNodes("x:brk[@man='1']", ns);
            var allBreakNodes = rowBreaksNode?.SelectNodes("x:brk", ns);
            var breakNodes = manualBreakNodes?.Count > 0 ? manualBreakNodes : allBreakNodes;

            if (breakNodes == null)
            {
                return ids;
            }

            foreach (XmlNode node in breakNodes)
            {
                var idText = node.Attributes?["id"]?.Value ?? string.Empty;
                if (int.TryParse(idText, out var id))
                {
                    ids.Add(id);
                }
            }

            return ids;
        }

        public static string FormatIdList(IEnumerable<int> ids)
        {
            return string.Join(", ", ids.OrderBy(x => x));
        }

        public static XmlNamespaceManager CreateWorksheetNamespaceManager(XmlDocument worksheetXml)
        {
            var ns = new XmlNamespaceManager(worksheetXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return ns;
        }

        private static bool TryGetDouble(object? value, out double number)
        {
            number = 0d;
            if (value == null)
            {
                return false;
            }

            switch (value)
            {
                case double d:
                    number = d;
                    return true;
                case float f:
                    number = f;
                    return true;
                case decimal dec:
                    number = (double)dec;
                    return true;
                case byte b:
                    number = b;
                    return true;
                case short s:
                    number = s;
                    return true;
                case int i:
                    number = i;
                    return true;
                case long l:
                    number = l;
                    return true;
                case string text:
                    return TryParseDouble(text, out number);
                default:
                    if (value is IConvertible convertible)
                    {
                        try
                        {
                            number = convertible.ToDouble(CultureInfo.InvariantCulture);
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }

                    return false;
            }
        }

        private static bool TryParseDouble(string? value, out double number)
        {
            number = 0d;
            var trimmed = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return false;
            }

            if (double.TryParse(trimmed, NumberStyles.Any, CultureInfo.InvariantCulture, out number))
            {
                return true;
            }

            return double.TryParse(trimmed, NumberStyles.Any, CultureInfo.GetCultureInfo("vi-VN"), out number);
        }
    }
}


