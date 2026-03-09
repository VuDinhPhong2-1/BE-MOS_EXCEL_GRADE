using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project06
{
    public class P06T4Grader : ITaskGrader
    {
        public string TaskId => "P06-T4";
        public string TaskName => "Summary B15 dung ham MAX cho cot Total sales";
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
                var ws = P06GraderHelpers.GetSheet(studentSheet, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Summary'.");
                    return result;
                }

                var cell = ws.Cells["B15"];
                var formulaRaw = cell.Formula ?? string.Empty;
                var formula = P06GraderHelpers.NormalizeFormula(formulaRaw);
                decimal score = 0;

                if (!string.IsNullOrWhiteSpace(formulaRaw))
                {
                    score += 1m;
                    result.Details.Add("O B15 da co cong thuc.");
                }
                else
                {
                    result.Errors.Add("O B15 chua co cong thuc.");
                    result.Score = score;
                    return result;
                }

                if (formula.Contains("MAX(", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("B15 da su dung ham MAX.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc B15 chua dung MAX. Hien tai: '{formulaRaw}'.");
                }

                var hasTotalSalesRef =
                    formula.Contains("TOTALSALES", StringComparison.Ordinal) ||
                    formula.Contains("F4:F11", StringComparison.Ordinal) ||
                    formula.Contains("F4:F12", StringComparison.Ordinal);
                if (hasTotalSalesRef)
                {
                    score += 1m;
                    result.Details.Add("Cong thuc B15 tham chieu dung cot Total sales.");
                }
                else
                {
                    result.Errors.Add("Cong thuc B15 chua tham chieu ro rang den cot Total sales.");
                }

                var table = ws.Tables.FirstOrDefault(t =>
                    t.Columns.Any(c => string.Equals((c.Name ?? string.Empty).Trim(), "Total sales", StringComparison.OrdinalIgnoreCase)));
                if (table != null)
                {
                    var tsCol = table.Columns.First(c =>
                        string.Equals((c.Name ?? string.Empty).Trim(), "Total sales", StringComparison.OrdinalIgnoreCase));
                    var colIdx = table.Address.Start.Column + tsCol.Position;
                    var startRow = table.Address.Start.Row + 1;
                    var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                    var maxValue = Enumerable.Range(startRow, endRow - startRow + 1)
                        .Select(r => P06GraderHelpers.ToDecimal(ws.Cells[r, colIdx].Value, ws.Cells[r, colIdx].Text))
                        .DefaultIfEmpty(0m)
                        .Max();

                    var b15Value = P06GraderHelpers.ToDecimal(cell.Value, cell.Text);
                    if (Math.Abs(maxValue - b15Value) <= 0.01m)
                    {
                        score += 1m;
                        result.Details.Add("Gia tri B15 khop gia tri lon nhat cua cot Total sales.");
                    }
                    else
                    {
                        result.Errors.Add($"Gia tri B15 ({b15Value}) khong khop MAX Total sales ({maxValue}).");
                    }
                }
                else
                {
                    result.Errors.Add("Khong tim thay table chua cot Total sales de doi chieu.");
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
