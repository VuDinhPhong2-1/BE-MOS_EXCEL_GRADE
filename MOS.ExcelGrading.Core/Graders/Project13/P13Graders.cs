using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project13
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

    public class P13T1Grader : ITaskGrader
    {
        public string TaskId => "P13-T1";
        public string TaskName => "Shirt Orders: replace Amber with Gold";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shirt Orders'.");
                    return result;
                }

                decimal score = 0m;
                var amberCount = 0;
                var goldCount = 0;
                for (var row = 6; row <= 199; row++)
                {
                    var color = (ws.Cells[row, 4].Text ?? string.Empty).Trim();
                    if (string.Equals(color, "Amber", StringComparison.OrdinalIgnoreCase))
                    {
                        amberCount++;
                    }

                    if (string.Equals(color, "Gold", StringComparison.Ordinal))
                    {
                        goldCount++;
                    }
                }

                if (amberCount == 0)
                {
                    score += 2m;
                    result.Details.Add("Khong con gia tri 'Amber' trong cot Shirt Color.");
                }
                else
                {
                    result.Errors.Add($"Van con {amberCount} gia tri 'Amber' trong cot Shirt Color.");
                }

                if (goldCount > 0)
                {
                    score += 2m;
                    result.Details.Add($"Da thay bang 'Gold' ({goldCount} o).");
                }
                else
                {
                    result.Errors.Add("Khong tim thay gia tri 'Gold' trong cot Shirt Color.");
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

    public class P13T2Grader : ITaskGrader
    {
        public string TaskId => "P13-T2";
        public string TaskName => "Shirt Orders: C2 formula SUMIF for Blue cost";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shirt Orders'.");
                    return result;
                }

                var isMatched = P13GraderHelpers.FormulaMatchesAny(
                    ws.Cells["C2"].Formula,
                    "SUMIF(D:D,\"Blue\",F:F)",
                    "SUMIF(D6:D199,\"Blue\",F6:F199)");

                if (isMatched)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc C2 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add(
                        $"Cong thuc C2 chua dung. Chap nhan SUMIF(D:D,\"Blue\",F:F) hoac SUMIF(D6:D199,\"Blue\",F6:F199). Hien tai: '{ws.Cells["C2"].Formula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P13T3Grader : ITaskGrader
    {
        public string TaskId => "P13-T3";
        public string TaskName => "Shirt Orders: C3 formula COUNTIF for Large";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shirt Orders'.");
                    return result;
                }

                var isMatched = P13GraderHelpers.FormulaMatchesAny(
                    ws.Cells["C3"].Formula,
                    "COUNTIF(E:E,\"Large\")",
                    "COUNTIF(E6:E199,\"Large\")");

                if (isMatched)
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc C3 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add(
                        $"Cong thuc C3 chua dung. Chap nhan COUNTIF(E:E,\"Large\") hoac COUNTIF(E6:E199,\"Large\"). Hien tai: '{ws.Cells["C3"].Formula}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }

    public class P13T4Grader : ITaskGrader
    {
        public string TaskId => "P13-T4";
        public string TaskName => "Shirt Orders: add subtotal row in D201/F201";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Shirt Orders'.");
                    return result;
                }

                decimal score = 0m;
                var studentD201 = ws.Cells["D201"];
                if (P13GraderHelpers.CellMatchesExpected(studentD201, "191"))
                {
                    score += 2m;
                    result.Details.Add("D201 dung gia tri mong doi (191).");
                }
                else
                {
                    result.Errors.Add(
                        $"D201 chua dung. Expected='191', Hien tai='{studentD201.Text}'.");
                }

                var studentF201 = ws.Cells["F201"];
                var actualFormula = P13GraderHelpers.NormalizeFormula(studentF201.Formula);
                if (string.IsNullOrWhiteSpace(actualFormula)
                    && string.IsNullOrWhiteSpace((studentF201.Text ?? string.Empty).Trim()))
                {
                    score += 2m;
                    result.Details.Add("F201 dung yeu cau de trong (khong co cong thuc, khong co gia tri).");
                }
                else
                {
                    result.Errors.Add(
                        $"F201 chua dung. Expected de trong, hien tai formula='{studentF201.Formula}', text='{studentF201.Text}'.");
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

    public class P13T5Grader : ITaskGrader
    {
        public string TaskId => "P13-T5";
        public string TaskName => "Attendees: page layout view + page break";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Attendees");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Attendees'.");
                    return result;
                }

                decimal score = 0m;
                if (ws.View.PageLayoutView)
                {
                    score += 2m;
                    result.Details.Add("Sheet dang o che do Page Layout.");
                }
                else
                {
                    result.Errors.Add("Sheet chua o che do Page Layout.");
                }

                var actualBreakIds = P13GraderHelpers.GetManualRowBreakIds(ws);
                var expectedBreakIds = new HashSet<int> { 35 };
                var hasExpectedBreak = actualBreakIds.SetEquals(expectedBreakIds);

                if (hasExpectedBreak)
                {
                    score += 2m;
                    result.Details.Add(
                        $"Da chen page break thu cong dung vi tri (id={P13GraderHelpers.FormatIdList(expectedBreakIds)}).");
                }
                else
                {
                    result.Errors.Add(
                        $"Page break thu cong chua dung. Expected id={P13GraderHelpers.FormatIdList(expectedBreakIds)}, actual id={P13GraderHelpers.FormatIdList(actualBreakIds)}.");
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

    public class P13T6Grader : ITaskGrader
    {
        public string TaskId => "P13-T6";
        public string TaskName => "Price List: H5 formula uses structured reference";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Price List");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Price List'.");
                    return result;
                }

                var actual = P13GraderHelpers.NormalizeFormula(ws.Cells["H5"].Formula);
                var expected = P13GraderHelpers.NormalizeFormula("ROWS(Phones[])");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc H5 dung: ROWS(Phones[]).");
                }
                else
                {
                    result.Errors.Add($"Cong thuc H5 chua dung. Hien tai: '{ws.Cells["H5"].Formula}'.");
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
