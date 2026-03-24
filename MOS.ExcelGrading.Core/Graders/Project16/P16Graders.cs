using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project16
{
    internal static class P16GraderHelpers
    {
        private static readonly Regex SqrefPartRegex = new(
            @"^(?<c1>[A-Z]+)(?<r1>\d+)(:(?<c2>[A-Z]+)(?<r2>\d+))?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

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
                .ToUpperInvariant();
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

        public static bool IsRangeMatch(string? actual, string expected)
        {
            return string.Equals(NormalizeRange(actual), NormalizeRange(expected), StringComparison.OrdinalIgnoreCase);
        }

        public static bool SqrefContainsRange(string? sqref, string expectedRange)
        {
            var normalizedExpected = NormalizeRange(expectedRange);
            var parts = (sqref ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return parts.Any(part =>
                string.Equals(NormalizeRange(part), normalizedExpected, StringComparison.OrdinalIgnoreCase));
        }

        public static string NormalizeIdentifier(string? value)
        {
            var text = (value ?? string.Empty).Trim().ToUpperInvariant();
            return new string(text.Where(char.IsLetterOrDigit).ToArray());
        }

        public static bool IsQuantityColumnName(string? value)
        {
            var normalized = NormalizeIdentifier(value);
            return normalized == "QUANTITY" || normalized == "QTY";
        }

        public static bool SqrefTargetsColumn(string? sqref, int expectedColumn, int dataStartRow, int dataEndRow)
        {
            var expectedColumnText = new string(ExcelCellBase.GetAddress(1, expectedColumn).Where(char.IsLetter).ToArray()).ToUpperInvariant();
            var parts = (sqref ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

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

                var row1 = int.Parse(match.Groups["r1"].Value);
                var row2 = int.Parse((match.Groups["r2"].Success ? match.Groups["r2"].Value : match.Groups["r1"].Value));
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

        public static (string ColorStyleId, string ChartStyleId) GetDrawingStyleIds(object drawing)
        {
            var styleManager = drawing.GetType().GetProperty("StyleManager")?.GetValue(drawing);
            if (styleManager == null)
            {
                return (string.Empty, string.Empty);
            }

            var colorsXml = styleManager.GetType().GetProperty("ColorsXml")?.GetValue(styleManager) as XmlDocument;
            var styleXml = styleManager.GetType().GetProperty("StyleXml")?.GetValue(styleManager) as XmlDocument;
            var colorStyleId = colorsXml?.DocumentElement?.Attributes?["id"]?.Value?.Trim() ?? string.Empty;
            var chartStyleId = styleXml?.DocumentElement?.Attributes?["id"]?.Value?.Trim() ?? string.Empty;
            return (colorStyleId, chartStyleId);
        }
    }

    public class P16T1Grader : ITaskGrader
    {
        public string TaskId => "P16-T1";
        public string TaskName => "Products: freeze top two rows";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Products'.");
                    return result;
                }

                decimal score = 0m;
                var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var pane = ws.WorksheetXml.SelectSingleNode("//x:sheetViews/x:sheetView/x:pane", ns);
                var ySplit = pane?.Attributes?["ySplit"]?.Value ?? string.Empty;
                var topLeft = pane?.Attributes?["topLeftCell"]?.Value ?? string.Empty;
                var state = pane?.Attributes?["state"]?.Value ?? string.Empty;

                if (string.Equals(ySplit, "2", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("ySplit dung = 2 (dong 1-2 duoc co dinh).");
                }
                else
                {
                    result.Errors.Add($"ySplit chua dung. Hien tai: '{ySplit}'.");
                }

                if (string.Equals(topLeft, "A3", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(state, "frozen", StringComparison.OrdinalIgnoreCase))
                {
                    score += 2m;
                    result.Details.Add("TopLeftCell va state dung (A3, frozen).");
                }
                else
                {
                    result.Errors.Add($"Trang thai pane chua dung. TopLeftCell='{topLeft}', State='{state}'.");
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

    public class P16T2Grader : ITaskGrader
    {
        public string TaskId => "P16-T2";
        public string TaskName => "Products: align A1 to left";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Products'.");
                    return result;
                }

                if (ws.Cells["A1"].Style.HorizontalAlignment == ExcelHorizontalAlignment.Left)
                {
                    result.Score = MaxScore;
                    result.Details.Add("A1 da duoc can trai.");
                }
                else
                {
                    result.Errors.Add($"A1 chua can trai. Hien tai: {ws.Cells["A1"].Style.HorizontalAlignment}.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P16T3Grader : ITaskGrader
    {
        public string TaskId => "P16-T3";
        public string TaskName => "Products: 3 traffic lights icon set on Quantity";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Products'.");
                    return result;
                }

                decimal score = 0m;
                var quantityCol = -1;
                var dataStart = -1;
                var dataEnd = -1;

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => P16GraderHelpers.IsQuantityColumnName(c.Name)));
                if (table != null)
                {
                    var quantityOffset = table.Columns
                        .Select((c, idx) => new { Column = c, Index = idx })
                        .First(x => P16GraderHelpers.IsQuantityColumnName(x.Column.Name))
                        .Index;
                    quantityCol = table.Address.Start.Column + quantityOffset;
                    dataStart = table.Address.Start.Row + 1;
                    dataEnd = table.Address.End.Row;
                }
                else if (P16GraderHelpers.TryFindColumnByHeader(ws, P16GraderHelpers.IsQuantityColumnName, out var headerRow, out var headerCol))
                {
                    quantityCol = headerCol;
                    dataStart = headerRow + 1;
                    dataEnd = P16GraderHelpers.GetLastDataRowInColumn(ws, headerCol, dataStart);
                }
                else
                {
                    result.Errors.Add("Khong tim thay cot 'Quantity'.");
                    return result;
                }

                if (dataStart <= 0 || dataEnd < dataStart)
                {
                    result.Errors.Add("Khong xac dinh duoc vung du lieu cot Quantity.");
                    return result;
                }

                var expectedRange = $"{ExcelCellBase.GetAddress(dataStart, quantityCol)}:{ExcelCellBase.GetAddress(dataEnd, quantityCol)}";

                var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var iconRules = ws.WorksheetXml.SelectNodes("//x:conditionalFormatting/x:cfRule[@type='iconSet']", ns);

                XmlNode? rule = null;
                if (iconRules != null)
                {
                    foreach (XmlNode candidate in iconRules)
                    {
                        var sqref = candidate.ParentNode?.Attributes?["sqref"]?.Value ?? string.Empty;
                        if (P16GraderHelpers.SqrefTargetsColumn(sqref, quantityCol, dataStart, dataEnd))
                        {
                            rule = candidate;
                            break;
                        }
                    }
                }

                if (rule != null)
                {
                    score += 2m;
                    result.Details.Add($"Tim thay icon set tren range {expectedRange}.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay icon set tren range {expectedRange}.");
                    result.Score = score;
                    return result;
                }

                var iconSetNode = rule.SelectSingleNode("x:iconSet", ns);
                var iconSetName = iconSetNode?.Attributes?["iconSet"]?.Value ?? string.Empty;
                var validIconSet = string.IsNullOrWhiteSpace(iconSetName)
                                   || string.Equals(iconSetName, "3TrafficLights1", StringComparison.OrdinalIgnoreCase);
                var cfvoNodes = iconSetNode?.SelectNodes("x:cfvo", ns);
                var valuesOk = cfvoNodes?.Count == 3
                               && string.Equals(cfvoNodes[0]?.Attributes?["type"]?.Value, "percent", StringComparison.OrdinalIgnoreCase)
                               && string.Equals(cfvoNodes[0]?.Attributes?["val"]?.Value, "0", StringComparison.Ordinal)
                               && string.Equals(cfvoNodes[1]?.Attributes?["val"]?.Value, "33", StringComparison.Ordinal)
                               && string.Equals(cfvoNodes[2]?.Attributes?["val"]?.Value, "67", StringComparison.Ordinal);

                if (validIconSet && valuesOk)
                {
                    score += 2m;
                    result.Details.Add("Icon set dung nguong 0/33/67 (3 traffic lights).");
                }
                else
                {
                    result.Errors.Add($"Icon set/nguong chua dung. iconSet='{iconSetName}', cfvoCount={cfvoNodes?.Count ?? 0}.");
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

    public class P16T4Grader : ITaskGrader
    {
        public string TaskId => "P16-T4";
        public string TaskName => "Products: table style Medium1";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Products'.");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault(t => P16GraderHelpers.IsRangeMatch(t.Address.Address, "A2:G54"));
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table A2:G54.");
                    return result;
                }

                if (table.TableStyle == TableStyles.Medium1)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Table style dung: TableStyleMedium1.");
                }
                else
                {
                    result.Errors.Add($"Table style chua dung. Hien tai: {table.TableStyle}.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P16T5Grader : ITaskGrader
    {
        public string TaskId => "P16-T5";
        public string TaskName => "Products: Estimated Value formula in F3 and fill down";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Products");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Products'.");
                    return result;
                }

                decimal score = 0m;
                var formulaF3 = P16GraderHelpers.NormalizeFormula(ws.Cells["F3"].Formula);
                var expectedF3 = P16GraderHelpers.NormalizeFormula("D3*E3");
                if (string.Equals(formulaF3, expectedF3, StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Cong thuc goc F3 dung: D3*E3.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc F3 chua dung. Hien tai: '{ws.Cells["F3"].Formula}'.");
                }

                var filledRows = 0;
                for (var row = 3; row <= 34; row++)
                {
                    if (!string.IsNullOrWhiteSpace(ws.Cells[row, 6].Formula))
                    {
                        filledRows++;
                    }
                }

                if (filledRows == 32)
                {
                    score += 2m;
                    result.Details.Add("Cong thuc da duoc fill down day du F3:F34.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc chua fill day du F3:F34 (hien tai {filledRows}/32).");
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

    public class P16T6Grader : ITaskGrader
    {
        public string TaskId => "P16-T6";
        public string TaskName => "Summary: chart Colorful Palette 2";
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
                var ws = P16GraderHelpers.GetSheet(studentSheet.Workbook, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Summary'.");
                    return result;
                }

                var drawing = ws.Drawings.FirstOrDefault();
                if (drawing == null)
                {
                    result.Errors.Add("Khong tim thay bieu do tren sheet 'Summary'.");
                    return result;
                }

                decimal score = 0m;
                score += 1m;
                result.Details.Add("Tim thay chart tren sheet 'Summary'.");

                var (colorStyleId, chartStyleId) = P16GraderHelpers.GetDrawingStyleIds(drawing);
                if (string.Equals(colorStyleId, "11", StringComparison.Ordinal))
                {
                    score += 2m;
                    result.Details.Add("Color style ID = 11 (Colorful Palette 2).");
                }
                else
                {
                    result.Errors.Add($"Color style ID chua dung. Hien tai: '{colorStyleId}' (mong doi 11).");
                }

                if (string.Equals(chartStyleId, "410", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Chart style ID dung: 410.");
                }
                else
                {
                    result.Errors.Add($"Chart style ID chua dung. Hien tai: '{chartStyleId}' (mong doi 410).");
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
}
