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
                   || (normalized.Contains("ORDER", StringComparison.Ordinal) && normalized.Contains("TOTAL", StringComparison.Ordinal));
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
                var orderTotalCol = -1;
                var dataStart = -1;
                var dataEnd = -1;

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => P15GraderHelpers.IsOrderTotalColumnName(c.Name)));
                if (table != null)
                {
                    var orderTotalOffset = table.Columns
                        .Select((c, idx) => new { Column = c, Index = idx })
                        .First(x => P15GraderHelpers.IsOrderTotalColumnName(x.Column.Name))
                        .Index;
                    orderTotalCol = table.Address.Start.Column + orderTotalOffset;
                    dataStart = table.Address.Start.Row + 1;
                    dataEnd = table.Address.End.Row;
                }
                else if (P15GraderHelpers.TryFindColumnByHeader(ws, P15GraderHelpers.IsOrderTotalColumnName, out var headerRow, out var headerCol))
                {
                    orderTotalCol = headerCol;
                    dataStart = headerRow + 1;
                    dataEnd = P15GraderHelpers.GetLastDataRowInColumn(ws, headerCol, dataStart);
                }
                else
                {
                    result.Errors.Add("Khong tim thay cot 'OrderTotal' tren sheet Orders.");
                    return result;
                }

                if (dataStart <= 0 || dataEnd < dataStart)
                {
                    result.Errors.Add("Khong xac dinh duoc vung du lieu cot OrderTotal.");
                    return result;
                }

                var expectedRange = $"{ExcelCellBase.GetAddress(dataStart, orderTotalCol)}:{ExcelCellBase.GetAddress(dataEnd, orderTotalCol)}";

                var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var aboveAverageRules = ws.WorksheetXml.SelectNodes(
                    "//x:conditionalFormatting/x:cfRule[@type='aboveAverage']",
                    ns);

                XmlNode? matchedRule = null;
                if (aboveAverageRules != null)
                {
                    foreach (XmlNode ruleNode in aboveAverageRules)
                    {
                        var sqref = ruleNode.ParentNode?.Attributes?["sqref"]?.Value ?? string.Empty;
                        if (P15GraderHelpers.SqrefTargetsColumn(sqref, orderTotalCol, dataStart, dataEnd))
                        {
                            matchedRule = ruleNode;
                            break;
                        }
                    }
                }

                if (matchedRule != null)
                {
                    score += 2m;
                    result.Details.Add($"Tim thay quy tac Above Average tren range {expectedRange}.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay quy tac Above Average tren range {expectedRange}.");
                    result.Score = score;
                    return result;
                }

                var dxfId = matchedRule.Attributes?["dxfId"]?.Value ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(dxfId))
                {
                    score += 2m;
                    result.Details.Add($"Rule co style dxfId={dxfId} (da ap dung style mau cho Above Average).");
                }
                else
                {
                    result.Errors.Add("Rule Above Average chua co dxfId de ap dung style.");
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

                decimal score = 0m;
                var filledCount = 0;
                var formatMatchedCount = 0;
                for (var row = 4; row <= 32; row++)
                {
                    var value = ws.Cells[row, 5].Text;
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        filledCount++;
                    }

                    if (ws.Cells[row, 5].StyleID == ws.Cells[row, 4].StyleID)
                    {
                        formatMatchedCount++;
                    }
                }

                if (filledCount == 29)
                {
                    score += 2m;
                    result.Details.Add("Cot CurrentAge da duoc dien day du (29/29 dong).");
                }
                else
                {
                    result.Errors.Add($"Cot CurrentAge chua dien du. Hien tai: {filledCount}/29 dong.");
                }

                if (formatMatchedCount == 29)
                {
                    score += 2m;
                    result.Details.Add("Dinh dang cot CurrentAge duoc giu nguyen.");
                }
                else
                {
                    result.Errors.Add($"Dinh dang cot CurrentAge bi thay doi o {29 - formatMatchedCount} dong.");
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
