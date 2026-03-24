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
                .ToUpperInvariant();
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

                var actual = P13GraderHelpers.NormalizeFormula(ws.Cells["C2"].Formula);
                var expected = P13GraderHelpers.NormalizeFormula("SUMIF(D:D,\"Blue\",F:F)");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc C2 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc C2 chua dung. Hien tai: '{ws.Cells["C2"].Formula}'.");
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

                var actual = P13GraderHelpers.NormalizeFormula(ws.Cells["C3"].Formula);
                var expected = P13GraderHelpers.NormalizeFormula("COUNTIF(E:E,\"Large\")");
                if (string.Equals(actual, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Cong thuc C3 dung chinh xac.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc C3 chua dung. Hien tai: '{ws.Cells["C3"].Formula}'.");
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
                var d201 = (ws.Cells["D201"].Text ?? string.Empty).Trim();
                if (string.Equals(d201, "Grand Total", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("D201 dung van ban 'Grand Total'.");
                }
                else
                {
                    result.Errors.Add($"D201 chua dung. Hien tai: '{d201}'.");
                }

                var formula = P13GraderHelpers.NormalizeFormula(ws.Cells["F201"].Formula);
                var expected = P13GraderHelpers.NormalizeFormula("SUBTOTAL(9,F6:F199)");
                if (string.Equals(formula, expected, StringComparison.Ordinal))
                {
                    score += 3m;
                    result.Details.Add("F201 dung cong thuc SUBTOTAL(9,F6:F199).");
                }
                else
                {
                    result.Errors.Add($"F201 chua dung cong thuc. Hien tai: '{ws.Cells["F201"].Formula}'.");
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

                var ns = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var breakNodes = ws.WorksheetXml.SelectNodes("//x:rowBreaks/x:brk[@man='1']", ns);
                var hasManualBreak = breakNodes?.Count > 0;

                var confirmedColumn = -1;
                for (var col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    var header = (ws.Cells[1, col].Text ?? string.Empty).Trim();
                    if (string.Equals(header, "Confirmed?", StringComparison.OrdinalIgnoreCase))
                    {
                        confirmedColumn = col;
                        break;
                    }
                }

                var expectedBreakId = -1;
                if (confirmedColumn > 0)
                {
                    var firstNonYRow = -1;
                    for (var row = 2; row <= ws.Dimension.End.Row; row++)
                    {
                        var value = (ws.Cells[row, confirmedColumn].Text ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            continue;
                        }

                        if (!string.Equals(value, "Y", StringComparison.OrdinalIgnoreCase))
                        {
                            firstNonYRow = row;
                            break;
                        }
                    }

                    if (firstNonYRow > 0)
                    {
                        expectedBreakId = firstNonYRow - 1;
                    }
                }

                var hasExpectedBreak = hasManualBreak
                    && expectedBreakId > 0
                    && breakNodes!.Cast<XmlNode>()
                        .Any(node => string.Equals(node.Attributes?["id"]?.Value, expectedBreakId.ToString(), StringComparison.Ordinal));

                if (hasExpectedBreak)
                {
                    score += 2m;
                    result.Details.Add($"Da chen page break thu cong dung vi tri (id={expectedBreakId}) de gom Y vao trang 1.");
                }
                else
                {
                    result.Errors.Add($"Page break thu cong chua dung. Expected id={expectedBreakId}, hasManualBreak={hasManualBreak}.");
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
