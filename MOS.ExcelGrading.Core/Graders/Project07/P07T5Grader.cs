using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project07
{
    public class P07T5Grader : ITaskGrader
    {
        public string TaskId => "P07-T5";
        public string TaskName => "Total Cookie Sales B3 dung SUM(Table2[Chocolate Mint Chip])";
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
                var ws = P07GraderHelpers.GetSheet(studentSheet, "Total Cookie Sales");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Total Cookie Sales'.");
                    return result;
                }

                var cell = ws.Cells["B3"];
                var formulaRaw = cell.Formula ?? string.Empty;
                var formula = P07GraderHelpers.NormalizeFormula(formulaRaw);
                decimal score = 0;

                if (!string.IsNullOrWhiteSpace(formulaRaw))
                {
                    score += 1m;
                    result.Details.Add("O B3 da co cong thuc.");
                }
                else
                {
                    result.Errors.Add("O B3 chua co cong thuc.");
                    result.Score = score;
                    return result;
                }

                if (formula.StartsWith("SUM(", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Cong thuc B3 da dung ham SUM.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc B3 chua dung SUM. Hien tai: '{formulaRaw}'.");
                }

                var usesStructuredRef =
                    formula.Contains("TABLE2[CHOCOLATEMINTCHIP]", StringComparison.Ordinal) ||
                    formula.Contains("TABLE2[[#DATA],[CHOCOLATEMINTCHIP]]", StringComparison.Ordinal);
                if (usesStructuredRef)
                {
                    score += 1m;
                    result.Details.Add("B3 da dung structured reference toi cot Chocolate Mint Chip.");
                }
                else
                {
                    result.Errors.Add("B3 chua dung structured reference den Table2[Chocolate Mint Chip].");
                }

                // Mot so file luu khong con object table nhung van giu structured-reference formula.
                // Khi do, kiem tra gia tri ket qua B3 la so hop le.
                if (decimal.TryParse(cell.Text, out _) || P07GraderHelpers.ToDecimal(cell.Value, cell.Text) > 0)
                {
                    score += 1m;
                    result.Details.Add("Gia tri B3 hop le sau khi tinh SUM.");
                }
                else
                {
                    result.Errors.Add("Gia tri B3 chua hop le sau khi tinh tong.");
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
