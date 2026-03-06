using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T4Grader : ITaskGrader
    {
        public string TaskId => "P02-T4";
        public string TaskName => "Dem so thang khong co policy moi trong cot Inactive months";
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
                var ws = studentSheet.Workbook.Worksheets["New Policy"];
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'New Policy'");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay bang du lieu New Policy");
                    return result;
                }

                var inactiveCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Inactive months", StringComparison.OrdinalIgnoreCase));
                if (inactiveCol == null)
                {
                    result.Errors.Add("Khong tim thay cot 'Inactive months'");
                    return result;
                }

                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var colIndex = table.Address.Start.Column + inactiveCol.Position;
                var totalRows = Math.Max(0, endRow - startRow + 1);

                if (totalRows == 0)
                {
                    result.Errors.Add("Khong co dong du lieu de cham");
                    return result;
                }

                var hasFormulaRows = 0;
                var countFunctionRows = 0;
                var janJunRefRows = 0;
                var blankConditionRows = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var formula = NormalizeFormula(ws.Cells[row, colIndex].Formula);
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        continue;
                    }

                    hasFormulaRows++;
                    if (formula.Contains("COUNTBLANK(") || formula.Contains("COUNTIF(") || formula.Contains("COUNTIFS("))
                    {
                        countFunctionRows++;
                    }

                    if (formula.Contains("JANUARY", StringComparison.OrdinalIgnoreCase) &&
                        formula.Contains("JUNE", StringComparison.OrdinalIgnoreCase))
                    {
                        janJunRefRows++;
                    }

                    if (formula.Contains("COUNTBLANK(") || formula.Contains("\"\""))
                    {
                        blankConditionRows++;
                    }
                }

                decimal score = 0;
                if (hasFormulaRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Cot Inactive months da co cong thuc cho tat ca dong");
                }
                else
                {
                    result.Errors.Add($"Thieu cong thuc tai cot Inactive months ({hasFormulaRows}/{totalRows})");
                }

                if (countFunctionRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Cong thuc co su dung ham dem");
                }
                else
                {
                    result.Errors.Add($"Ham dem chua dung/du ({countFunctionRows}/{totalRows})");
                }

                if (janJunRefRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Cong thuc tham chieu dung khoang January:June");
                }
                else
                {
                    result.Errors.Add($"Tham chieu January:June chua dung/du ({janJunRefRows}/{totalRows})");
                }

                if (blankConditionRows == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Cong thuc co dieu kien dem thang trong");
                }
                else
                {
                    result.Errors.Add($"Dieu kien dem o trong chua day du ({blankConditionRows}/{totalRows})");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static string NormalizeFormula(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return string.Empty;
            }

            return formula
                .Replace("=", string.Empty)
                .Replace("$", string.Empty)
                .Replace(" ", string.Empty)
                .ToUpperInvariant();
        }
    }
}

