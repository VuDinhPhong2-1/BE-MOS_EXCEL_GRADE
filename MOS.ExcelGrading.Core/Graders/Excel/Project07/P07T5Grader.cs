using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project07
{
    public class P07T5Grader : ITaskGrader
    {
        public string TaskId => "P07-T5";
        public string TaskName => "Total Cookie Sales B3 dùng SUM(Table2[Chocolate Mint Chip])";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet)
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
                    result.Errors.Add("Không tìm thấy sheet 'Total Cookie Sales'.");
                    return result;
                }

                var cell = ws.Cells["B3"];
                var formulaRaw = cell.Formula ?? string.Empty;
                var formula = P07GraderHelpers.NormalizeFormula(formulaRaw);
                decimal score = 0;

                if (!string.IsNullOrWhiteSpace(formulaRaw))
                {
                    score += 1m;
                    result.Details.Add("Ô B3 đã có công thức.");
                }
                else
                {
                    result.Errors.Add("Ô B3 chưa có công thức.");
                    result.Score = score;
                    return result;
                }

                if (formula.StartsWith("SUM(", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Công thức B3 đã dùng hàm SUM.");
                }
                else
                {
                    result.Errors.Add($"Công thức B3 chưa dùng SUM. Hiện tại: '{formulaRaw}'.");
                }

                var usesStructuredRef =
                    formula.Contains("TABLE2[CHOCOLATEMINTCHIP]", StringComparison.Ordinal) ||
                    formula.Contains("TABLE2[[#DATA],[CHOCOLATEMINTCHIP]]", StringComparison.Ordinal);
                if (usesStructuredRef)
                {
                    score += 1m;
                    result.Details.Add("B3 đã dùng structured reference tới cột Chocolate Mint Chip.");
                }
                else
                {
                    result.Errors.Add("B3 chưa dùng structured reference đến Table2[Chocolate Mint Chip].");
                }

                // Mot so file luu khong con object table nhung van giu structured-reference formula.
                // Khi do, kiem tra gia tri ket qua B3 la so hop le.
                if (decimal.TryParse(cell.Text, out _) || P07GraderHelpers.ToDecimal(cell.Value, cell.Text) > 0)
                {
                    score += 1m;
                    result.Details.Add("Giá trị B3 hợp lệ sau khi tính SUM.");
                }
                else
                {
                    result.Errors.Add("Giá trị B3 chưa hợp lệ sau khi tính tổng.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}

// minor-sync: non-functional graders update
