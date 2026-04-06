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
                return "Khong doc duoc style XML.";
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

    public class P15T1Grader : ITaskGrader
    {
        public string TaskId => "P15-T1";
        public string TaskName => "Products: Weight number format with 3 decimals";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Products'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => string.Equals(c.Name, "Weight", StringComparison.OrdinalIgnoreCase)));
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table co cot 'Weight'.");
                    return result;
                }

                var weightOffset = table.Columns
                    .Select((c, idx) => new { Column = c, Index = idx })
                    .First(x => string.Equals(x.Column.Name, "Weight", StringComparison.OrdinalIgnoreCase))
                    .Index;
                var weightCol = table.Address.Start.Column + weightOffset;

                var dataStart = table.Address.Start.Row + 1;
                var dataEnd = table.Address.End.Row;
                var totalRows = Math.Max(0, dataEnd - dataStart + 1);
                var validRows = 0;
                for (var row = dataStart; row <= dataEnd; row++)
                {
                    var format = ws.Cells[row, weightCol].Style.Numberformat.Format;
                    if (P15GraderHelpers.IsThreeDecimalNumberFormat(format))
                    {
                        validRows++;
                    }
                }

                if (validRows == totalRows && totalRows > 0)
                {
                    result.Score = MaxScore;
                    result.Details.Add($"Dinh dang 3 so thap phan dung tren toan bo cot Weight ({validRows}/{totalRows}).");
                }
                else
                {
                    result.Score = 2m;
                    result.Errors.Add($"Dinh dang cot Weight chua dung tren toan bo du lieu ({validRows}/{totalRows}).");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P15T2Grader : ITaskGrader
    {
        public string TaskId => "P15-T2";
        public string TaskName => "Products: G3 formula SUMIF for Magic Supplies";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Products'.");
                    return result;
                }

                var actual = P15GraderHelpers.NormalizeFormula(ws.Cells["G3"].Formula);
                var expected = P15GraderHelpers.NormalizeFormula("SUMIF(Table2[Catergory],\"Magic Supplies\",Table2[Weight])");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc G3 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc G3 chua dung. Hien tai: '{ws.Cells["G3"].Formula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P15T3Grader : ITaskGrader
    {
        public string TaskId => "P15-T3";
        public string TaskName => "Orders: above average conditional format with green style";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Orders");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Orders'.");
                    return result;
                }

                decimal score = 0m;
                if (!P15GraderHelpers.TryFindOrderTotalDataRange(ws, out var orderTotalCol, out var dataStart, out var dataEnd))
                {
                    result.Errors.Add("Khong tim thay cot 'OrderTotal' hoac khong xac dinh duoc vung du lieu tren sheet Orders.");
                    return result;
                }

                var expectedRange = $"{ExcelCellBase.GetAddress(dataStart, orderTotalCol)}:{ExcelCellBase.GetAddress(dataEnd, orderTotalCol)}";
                var targetedRules = ws.ConditionalFormatting
                    .Where(cf => P15GraderHelpers.SqrefTargetsColumn(cf.Address.Address, orderTotalCol, dataStart, dataEnd))
                    .ToList();

                var matchedRule = targetedRules.FirstOrDefault(cf =>
                    cf.Type == OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.AboveAverage);

                if (matchedRule == null)
                {
                    matchedRule = targetedRules.FirstOrDefault(cf =>
                    {
                        if (cf.Type != OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.Expression
                            && cf.Type != OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.GreaterThan)
                        {
                            return false;
                        }

                        var formula = cf.GetType().GetProperty("Formula")?.GetValue(cf)?.ToString() ?? string.Empty;
                        return P15GraderHelpers.IsFormulaBasedAboveAverageRule(formula, orderTotalCol, dataStart, dataEnd);
                    });
                }

                if (matchedRule == null)
                {
                    if (targetedRules.Count == 0)
                    {
                        result.Errors.Add($"Khong tim thay conditional formatting rule tren range {expectedRange}.");
                    }
                    else
                    {
                        var types = string.Join(", ", targetedRules.Select(r => r.Type.ToString()).Distinct(StringComparer.OrdinalIgnoreCase));
                        result.Errors.Add($"Co rule tren range {expectedRange} nhung khong phai quy tac 'Above Average' (types: {types}).");
                    }

                    result.Score = score;
                    return result;
                }

                score += 2m;
                result.Details.Add($"Tim thay quy tac Above Average (dong) tren range {expectedRange}.");

                var fontColor = matchedRule.Style.Font.Color.Color;
                var fillBgColor = matchedRule.Style.Fill.BackgroundColor.Color;
                var fillPatternColor = matchedRule.Style.Fill.PatternColor.Color;

                var fontHex = fontColor.HasValue ? P15GraderHelpers.ToArgbHex(fontColor.Value.ToArgb()) : string.Empty;
                var bgHex = fillBgColor.HasValue ? P15GraderHelpers.ToArgbHex(fillBgColor.Value.ToArgb()) : string.Empty;
                var patternHex = fillPatternColor.HasValue ? P15GraderHelpers.ToArgbHex(fillPatternColor.Value.ToArgb()) : string.Empty;

                var isGreenFill = P15GraderHelpers.IsGreenFillColor(bgHex) || P15GraderHelpers.IsGreenFillColor(patternHex);
                var isDarkGreenText = P15GraderHelpers.IsDarkGreenTextColor(fontHex);
                if (isGreenFill && isDarkGreenText)
                {
                    score += 2m;
                    result.Details.Add("Style quy tac dung yeu cau: Green Fill with Dark Green Text.");
                }
                else
                {
                    result.Errors.Add($"Style quy tac chua dung Green Fill with Dark Green Text. fill.bg={bgHex}, fill.pattern={patternHex}, font={fontHex}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P15T4Grader : ITaskGrader
    {
        public string TaskId => "P15-T4";
        public string TaskName => "Customers: N5 formula COUNTIF United States";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Customers");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Customers'.");
                    return result;
                }

                var actual = P15GraderHelpers.NormalizeFormula(ws.Cells["N5"].Formula);
                var expected = P15GraderHelpers.NormalizeFormula("COUNTIF(I4:I32,\"United States\")");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc N5 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc N5 chua dung. Hien tai: '{ws.Cells["N5"].Formula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P15T5Grader : ITaskGrader
    {
        public string TaskId => "P15-T5";
        public string TaskName => "Customers: complete CurrentAge column without format changes";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Customers");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Customers'.");
                    return result;
                }

                if (!P15GraderHelpers.TryFindColumnByHeader(
                    ws,
                    P15GraderHelpers.IsCurrentAgeColumnName,
                    out var headerRow,
                    out var currentAgeCol))
                {
                    result.Errors.Add("Khong tim thay cot 'CurrentAge'.");
                    return result;
                }

                var dataStart = headerRow + 1;
                var dataEnd = ws.Dimension.End.Row;
                if (P15GraderHelpers.TryFindColumnByHeader(ws, P15GraderHelpers.IsIdColumnName, out _, out var idCol))
                {
                    dataEnd = P15GraderHelpers.GetLastDataRowInColumn(ws, idCol, dataStart);
                }

                if (dataStart <= 0 || dataEnd < dataStart)
                {
                    result.Errors.Add("Khong xac dinh duoc vung du lieu cot CurrentAge.");
                    return result;
                }

                var totalRows = dataEnd - dataStart + 1;
                var hasBirthDateCol = P15GraderHelpers.TryFindColumnByHeader(
                    ws,
                    P15GraderHelpers.IsBirthDateColumnName,
                    out _,
                    out var birthDateCol);

                decimal score = 0m;
                var filledCount = 0;
                var seriesCorrectCount = 0;
                var styleIds = new List<int>(capacity: totalRows);

                for (var row = dataStart; row <= dataEnd; row++)
                {
                    var ageCell = ws.Cells[row, currentAgeCol];
                    styleIds.Add(ageCell.StyleID);

                    var hasFormulaProp = !string.IsNullOrWhiteSpace(ageCell.Formula);
                    var hasFormulaNode = P15GraderHelpers.TryGetCellFormulaNode(ws, row, currentAgeCol, out var formulaNode);
                    var hasValue = !string.IsNullOrWhiteSpace(ageCell.Text);
                    if (hasValue || hasFormulaProp || hasFormulaNode)
                    {
                        filledCount++;
                    }

                    var formulaText = hasFormulaProp
                        ? ageCell.Formula
                        : (hasFormulaNode ? (formulaNode?.InnerText ?? string.Empty) : string.Empty);
                    var normalizedFormula = P15GraderHelpers.NormalizeFormula(formulaText);
                    var isFormulaRow = hasFormulaProp || hasFormulaNode;

                    var isSeriesCorrect = false;
                    if (isFormulaRow)
                    {
                        if (string.IsNullOrWhiteSpace(normalizedFormula))
                        {
                            // Shared-formula follower cells can have empty inner text in XML.
                            isSeriesCorrect = true;
                        }
                        else if (hasBirthDateCol)
                        {
                            var birthCellAddress = ExcelCellBase.GetAddress(row, birthDateCol);
                            var birthRef = P15GraderHelpers.NormalizeFormula(birthCellAddress);
                            isSeriesCorrect =
                                normalizedFormula.Contains("YEAR(TODAY())", StringComparison.Ordinal)
                                && normalizedFormula.Contains($"YEAR({birthRef})", StringComparison.Ordinal);
                        }
                    }
                    else if (hasBirthDateCol && hasValue)
                    {
                        // Accept pre-calculated values if they match YEAR(TODAY())-YEAR(BirthDate).
                        var birthValue = ws.Cells[row, birthDateCol].Value;
                        DateTime birthDate;
                        if (birthValue is DateTime dt)
                        {
                            birthDate = dt;
                        }
                        else if (birthValue is double oa)
                        {
                            birthDate = DateTime.FromOADate(oa);
                        }
                        else if (!DateTime.TryParse(ws.Cells[row, birthDateCol].Text, out birthDate))
                        {
                            birthDate = DateTime.MinValue;
                        }

                        if (birthDate != DateTime.MinValue && int.TryParse(ageCell.Text, out var ageValue))
                        {
                            var expectedAge = DateTime.Today.Year - birthDate.Year;
                            isSeriesCorrect = ageValue == expectedAge;
                        }
                    }

                    if (isSeriesCorrect)
                    {
                        seriesCorrectCount++;
                    }
                }

                if (filledCount == totalRows)
                {
                    score += 2m;
                    result.Details.Add($"Cot CurrentAge da duoc dien day du ({filledCount}/{totalRows} dong).");
                }
                else
                {
                    result.Errors.Add($"Cot CurrentAge chua dien du. Hien tai: {filledCount}/{totalRows} dong.");
                }

                if (seriesCorrectCount == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Chuoi du lieu CurrentAge dung (cong thuc/ket qua hop le tren tat ca dong).");
                }
                else
                {
                    result.Errors.Add($"Chuoi du lieu CurrentAge chua dung ({seriesCorrectCount}/{totalRows} dong hop le).");
                }

                var innerRows = Enumerable.Range(dataStart + 1, Math.Max(0, totalRows - 2)).ToList();
                var oddStyles = innerRows
                    .Where(r => (r - dataStart) % 2 == 1)
                    .Select(r => ws.Cells[r, currentAgeCol].StyleID)
                    .Distinct()
                    .ToList();
                var evenStyles = innerRows
                    .Where(r => (r - dataStart) % 2 == 0)
                    .Select(r => ws.Cells[r, currentAgeCol].StyleID)
                    .Distinct()
                    .ToList();
                var styleOk = styleIds.All(id => id > 0)
                              && oddStyles.Count <= 1
                              && evenStyles.Count <= 1
                              && (oddStyles.Count == 0 || evenStyles.Count == 0 || oddStyles[0] != evenStyles[0]);
                if (styleOk)
                {
                    score += 1m;
                    result.Details.Add("Dinh dang cot CurrentAge duoc giu on dinh theo mau dong.");
                }
                else
                {
                    result.Errors.Add("Dinh dang cot CurrentAge co dau hieu bi thay doi (khong con on dinh theo mau dong).");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P15T6Grader : ITaskGrader
    {
        public string TaskId => "P15-T6";
        public string TaskName => "Set worksheet tab color Pink Accent 1";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var targetSheets = new[] { "Customers", "Products", "Orders" };
                var matched = 0;
                foreach (var name in targetSheets)
                {
                    var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, name);
                    if (ws == null)
                    {
                        continue;
                    }

                    var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                    ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    var tabColor = ws.WorksheetXml.SelectSingleNode("//x:sheetPr/x:tabColor", ns);
                    var theme = tabColor?.Attributes?["theme"]?.Value ?? string.Empty;
                    if (string.Equals(theme, "4", StringComparison.Ordinal))
                    {
                        matched++;
                    }
                }

                if (matched == targetSheets.Length)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Da dat mau tab Pink Accent 1 cho Customers, Products, Orders.");
                }
                else
                {
                    result.Score = 2m;
                    result.Errors.Add($"Mau tab chua dung day du ({matched}/{targetSheets.Length} sheet).");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
