using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project15
{
    internal static class P15GraderHelpers
    {
        private static readonly Regex SqrefPartRegex = new(
            @"^(?<c1>[A-Z]+)(?<r1>\d+)?(?::(?<c2>[A-Z]+)(?<r2>\d+)?)?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex AverageArgRegex = new(
            @"AVERAGE\((?<args>[^)]*)\)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly HashSet<string> GreenFillColors = new(StringComparer.OrdinalIgnoreCase)
        {
            "FFC6EFCE",
            "C6EFCE"
        };
        private static readonly HashSet<string> DarkGreenTextColors = new(StringComparer.OrdinalIgnoreCase)
        {
            "FF006100",
            "006100"
        };

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

        public static bool IsThreeDecimalNumberFormat(string? format)
        {
            var value = (format ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var normalized = value.Replace(" ", string.Empty, StringComparison.Ordinal);
            if (normalized.Contains("0.0000", StringComparison.Ordinal))
            {
                return false;
            }

            return normalized.Contains("0.000", StringComparison.Ordinal);
        }

        public static string NormalizeRange(string? value)
        {
            return (value ?? string.Empty)
                .Trim()
                .Replace("$", string.Empty, StringComparison.Ordinal)
                .Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();
        }

        public static bool SqrefContainsRange(string? sqref, string expectedRange)
        {
            var normalizedExpected = NormalizeRange(expectedRange);
            var parts = (sqref ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            return parts.Any(part => string.Equals(NormalizeRange(part), normalizedExpected, StringComparison.Ordinal));
        }

        public static string NormalizeIdentifier(string? value)
        {
            var text = (value ?? string.Empty).Trim().ToUpperInvariant();
            return new string(text.Where(char.IsLetterOrDigit).ToArray());
        }

        public static bool IsOrderTotalColumnName(string? value)
        {
            var normalized = NormalizeIdentifier(value);
            return normalized == "ORDERTOTAL"
                   || normalized == "ODERTOTAL"
                   || (normalized.Contains("ORDER", StringComparison.Ordinal) && normalized.Contains("TOTAL", StringComparison.Ordinal));
        }

        public static bool IsCurrentAgeColumnName(string? value)
        {
            return string.Equals(NormalizeIdentifier(value), "CURRENTAGE", StringComparison.Ordinal);
        }

        public static bool IsBirthDateColumnName(string? value)
        {
            return string.Equals(NormalizeIdentifier(value), "BIRTHDATE", StringComparison.Ordinal);
        }

        public static bool IsIdColumnName(string? value)
        {
            return string.Equals(NormalizeIdentifier(value), "ID", StringComparison.Ordinal);
        }

        public static bool SqrefTargetsColumn(string? sqref, int expectedColumn, int dataStartRow, int dataEndRow)
        {
            var expectedColumnText = new string(ExcelCellBase.GetAddress(1, expectedColumn).Where(char.IsLetter).ToArray()).ToUpperInvariant();
            var parts = (sqref ?? string.Empty)
                .Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            foreach (var part in parts)
            {
                var normalizedPart = NormalizeRange(part);
                var match = SqrefPartRegex.Match(normalizedPart);
                if (!match.Success)
                {
                    continue;
                }

                var col1 = match.Groups["c1"].Value.ToUpperInvariant();
                var col2 = (match.Groups["c2"].Success ? match.Groups["c2"].Value : match.Groups["c1"].Value).ToUpperInvariant();
                if (!string.Equals(col1, expectedColumnText, StringComparison.Ordinal)
                    || !string.Equals(col2, expectedColumnText, StringComparison.Ordinal))
                {
                    continue;
                }

                var hasCol2 = match.Groups["c2"].Success;
                var hasRow1 = int.TryParse(match.Groups["r1"].Value, out var parsedRow1);
                var hasRow2 = int.TryParse(match.Groups["r2"].Value, out var parsedRow2);

                // Whole-column reference (e.g., G:G or G).
                if (!hasRow1 && !hasRow2)
                {
                    return true;
                }

                var row1 = dataStartRow;
                var row2 = dataEndRow;

                if (hasCol2)
                {
                    row1 = hasRow1 ? parsedRow1 : dataStartRow;
                    row2 = hasRow2 ? parsedRow2 : dataEndRow;
                }
                else
                {
                    if (!hasRow1)
                    {
                        return true;
                    }

                    row1 = parsedRow1;
                    row2 = parsedRow1;
                }

                var minRow = Math.Min(row1, row2);
                var maxRow = Math.Max(row1, row2);

                if (maxRow >= dataStartRow && minRow <= dataEndRow)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool TryFindColumnByHeader(
            ExcelWorksheet worksheet,
            Func<string, bool> headerMatcher,
            out int headerRow,
            out int columnIndex)
        {
            headerRow = -1;
            columnIndex = -1;

            var maxHeaderRow = Math.Min(10, worksheet.Dimension.End.Row);
            for (var row = 1; row <= maxHeaderRow; row++)
            {
                for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var text = (worksheet.Cells[row, col].Text ?? string.Empty).Trim();
                    if (headerMatcher(text))
                    {
                        headerRow = row;
                        columnIndex = col;
                        return true;
                    }
                }
            }

            return false;
        }

        public static int GetLastDataRowInColumn(ExcelWorksheet worksheet, int columnIndex, int startRow)
        {
            var lastRow = startRow;
            for (var row = startRow; row <= worksheet.Dimension.End.Row; row++)
            {
                var cell = worksheet.Cells[row, columnIndex];
                if (!string.IsNullOrWhiteSpace(cell.Text) || !string.IsNullOrWhiteSpace(cell.Formula))
                {
                    lastRow = row;
                }
            }

            return lastRow;
        }

        public static bool TryGetCellFormulaNode(ExcelWorksheet worksheet, int row, int column, out XmlNode? formulaNode)
        {
            formulaNode = null;
            var xml = worksheet.WorksheetXml;
            if (xml == null)
            {
                return false;
            }

            var ns = new XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var cellAddress = ExcelCellBase.GetAddress(row, column);
            formulaNode = xml.SelectSingleNode(
                $"/x:worksheet/x:sheetData/x:row[@r='{row}']/x:c[@r='{cellAddress}']/x:f",
                ns);
            return formulaNode != null;
        }

        public static bool TryFindOrderTotalDataRange(
            ExcelWorksheet worksheet,
            out int orderTotalColumn,
            out int dataStartRow,
            out int dataEndRow)
        {
            orderTotalColumn = -1;
            dataStartRow = -1;
            dataEndRow = -1;

            var table = worksheet.Tables.FirstOrDefault(t =>
                t.Columns.Any(c => IsOrderTotalColumnName(c.Name)));
            if (table != null)
            {
                var orderTotalOffset = table.Columns
                    .Select((c, idx) => new { Column = c, Index = idx })
                    .First(x => IsOrderTotalColumnName(x.Column.Name))
                    .Index;
                orderTotalColumn = table.Address.Start.Column + orderTotalOffset;
                dataStartRow = table.Address.Start.Row + 1;
                dataEndRow = table.Address.End.Row;
                return dataStartRow > 0 && dataEndRow >= dataStartRow;
            }

            if (TryFindColumnByHeader(worksheet, IsOrderTotalColumnName, out var headerRow, out var headerCol))
            {
                orderTotalColumn = headerCol;
                dataStartRow = headerRow + 1;
                dataEndRow = GetLastDataRowInColumn(worksheet, headerCol, dataStartRow);
                return dataStartRow > 0 && dataEndRow >= dataStartRow;
            }

            return false;
        }

        public static bool IsAboveAverageConditionRule(
            XmlNode ruleNode,
            int expectedColumn,
            int dataStartRow,
            int dataEndRow)
        {
            var type = (ruleNode.Attributes?["type"]?.Value ?? string.Empty).Trim();
            if (string.Equals(type, "aboveAverage", StringComparison.OrdinalIgnoreCase))
            {
                var aboveAverage = ruleNode.Attributes?["aboveAverage"]?.Value ?? "1";
                var equalAverage = ruleNode.Attributes?["equalAverage"]?.Value ?? "0";
                return !string.Equals(aboveAverage, "0", StringComparison.Ordinal)
                    && !string.Equals(equalAverage, "1", StringComparison.Ordinal);
            }

            if (!string.Equals(type, "expression", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(type, "cellIs", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.Equals(type, "cellIs", StringComparison.OrdinalIgnoreCase))
            {
                var op = ruleNode.Attributes?["operator"]?.Value ?? string.Empty;
                if (!string.Equals(op, "greaterThan", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            var doc = ruleNode.OwnerDocument;
            if (doc == null)
            {
                return false;
            }

            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var formulaNodes = ruleNode.SelectNodes("x:formula", ns);
            if (formulaNodes == null || formulaNodes.Count == 0)
            {
                return false;
            }

            foreach (XmlNode formulaNode in formulaNodes)
            {
                if (IsFormulaBasedAboveAverage(formulaNode.InnerText, expectedColumn, dataStartRow, dataEndRow))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool TryGetRuleDxfNode(ExcelWorkbook workbook, XmlNode ruleNode, out XmlNode? dxfNode)
        {
            dxfNode = null;
            var dxfIdRaw = ruleNode.Attributes?["dxfId"]?.Value;
            if (!int.TryParse(dxfIdRaw, out var dxfId) || dxfId < 0)
            {
                return false;
            }

            var stylesXml = workbook.StylesXml;
            if (stylesXml == null)
            {
                return false;
            }

            var ns = new XmlNamespaceManager(stylesXml.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            dxfNode = stylesXml.SelectSingleNode($"/x:styleSheet/x:dxfs/x:dxf[{dxfId + 1}]", ns);
            return dxfNode != null;
        }

        public static bool IsGreenFillWithDarkGreenText(XmlNode dxfNode)
        {
            var doc = dxfNode.OwnerDocument;
            if (doc == null)
            {
                return false;
            }

            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var fgColorNode = dxfNode.SelectSingleNode("x:fill/x:patternFill/x:fgColor", ns);
            var bgColorNode = dxfNode.SelectSingleNode("x:fill/x:patternFill/x:bgColor", ns);
            var fontColorNode = dxfNode.SelectSingleNode("x:font/x:color", ns);

            var hasGreenFill = IsKnownColor(fgColorNode, GreenFillColors) || IsKnownColor(bgColorNode, GreenFillColors);
            var hasDarkGreenText = IsKnownColor(fontColorNode, DarkGreenTextColors);
            return hasGreenFill && hasDarkGreenText;
        }

        public static bool IsFormulaBasedAboveAverageRule(string? formula, int expectedColumn, int dataStartRow, int dataEndRow)
        {
            return IsFormulaBasedAboveAverage(formula, expectedColumn, dataStartRow, dataEndRow);
        }

        public static bool IsGreenFillColor(string? color)
        {
            var normalized = NormalizeRgb(color);
            return !string.IsNullOrWhiteSpace(normalized) && GreenFillColors.Contains(normalized);
        }

        public static bool IsDarkGreenTextColor(string? color)
        {
            var normalized = NormalizeRgb(color);
            return !string.IsNullOrWhiteSpace(normalized) && DarkGreenTextColors.Contains(normalized);
        }

        public static string ToArgbHex(int argb)
        {
            return unchecked((uint)argb).ToString("X8");
        }

        public static string DescribeDxfColors(XmlNode dxfNode)
        {
            var doc = dxfNode.OwnerDocument;
            if (doc == null)
            {
                return "Không đọc được style XML.";
            }

            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var fg = dxfNode.SelectSingleNode("x:fill/x:patternFill/x:fgColor", ns);
            var bg = dxfNode.SelectSingleNode("x:fill/x:patternFill/x:bgColor", ns);
            var font = dxfNode.SelectSingleNode("x:font/x:color", ns);

            return $"fill.fg={GetColorDebugText(fg)}, fill.bg={GetColorDebugText(bg)}, font={GetColorDebugText(font)}";
        }

        private static bool IsFormulaBasedAboveAverage(string? formula, int expectedColumn, int dataStartRow, int dataEndRow)
        {
            var normalized = NormalizeFormula(formula);
            if (string.IsNullOrWhiteSpace(normalized) || !normalized.Contains("AVERAGE(", StringComparison.Ordinal))
            {
                return false;
            }

            var matches = AverageArgRegex.Matches(normalized);
            foreach (Match match in matches)
            {
                var argsText = match.Groups["args"].Value;
                var args = argsText.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (var arg in args)
                {
                    if (SqrefTargetsColumn(arg, expectedColumn, dataStartRow, dataEndRow))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsKnownColor(XmlNode? colorNode, HashSet<string> acceptedColors)
        {
            if (colorNode == null)
            {
                return false;
            }

            var rgb = NormalizeRgb(colorNode.Attributes?["rgb"]?.Value);
            return !string.IsNullOrWhiteSpace(rgb) && acceptedColors.Contains(rgb);
        }

        private static string NormalizeRgb(string? rgb)
        {
            return (rgb ?? string.Empty)
                .Trim()
                .Replace("#", string.Empty, StringComparison.Ordinal)
                .ToUpperInvariant();
        }

        private static string GetColorDebugText(XmlNode? colorNode)
        {
            if (colorNode == null)
            {
                return "(none)";
            }

            var rgb = NormalizeRgb(colorNode.Attributes?["rgb"]?.Value);
            var theme = colorNode.Attributes?["theme"]?.Value ?? string.Empty;
            var tint = colorNode.Attributes?["tint"]?.Value ?? string.Empty;
            var indexed = colorNode.Attributes?["indexed"]?.Value ?? string.Empty;
            return $"rgb={rgb};theme={theme};tint={tint};indexed={indexed}";
        }
    }
}

