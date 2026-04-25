using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T3Grader : ITaskGrader
    {
        public string TaskId => "P05-T3";
        public string TaskName => "F37 tính trung bình Selling Price cho 'Fabrikam, Inc.'";
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
                var ws = P05GraderHelpers.GetSheet(studentSheet, "Annual Purchases");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Annual Purchases'.");
                    return result;
                }

                var cell = ws.Cells["F37"];
                var formulaRaw = cell.Formula;
                var formula = P05GraderHelpers.NormalizeFormula(formulaRaw);

                decimal score = 0;
                if (!string.IsNullOrWhiteSpace(formulaRaw))
                {
                    score += 1m;
                    result.Details.Add("Ô F37 đã có công thức.");
                }
                else
                {
                    result.Errors.Add("Ô F37 chưa có công thức.");
                    result.Score = score;
                    return result;
                }

                var usesAverage = formula.StartsWith("AVERAGEIF(", StringComparison.Ordinal) ||
                                  formula.StartsWith("AVERAGEIFS(", StringComparison.Ordinal);
                if (usesAverage)
                {
                    score += 1m;
                    result.Details.Add("Công thức đã dùng hàm AVERAGEIF/AVERAGEIFS.");
                }
                else
                {
                    result.Errors.Add($"Công thức F37 chưa dùng hàm AVERAGEIF. Hiện tại: '{formulaRaw}'.");
                }

                var hasPublisherRange = formula.Contains("D5:D35", StringComparison.Ordinal);
                var hasSellingRange = formula.Contains("F5:F35", StringComparison.Ordinal);
                if (hasPublisherRange && hasSellingRange)
                {
                    score += 1m;
                    result.Details.Add("Công thức tham chiếu đúng các vùng D5:D35 và F5:F35.");
                }
                else
                {
                    result.Errors.Add(
                        $"Công thức F37 chưa tham chiếu đúng range (cần có D5:D35 và F5:F35). Formula: '{formulaRaw}'.");
                }

                var hasPublisherCriteria = formula.Contains("\"FABRIKAM,INC.\"", StringComparison.Ordinal) ||
                                           formula.Contains("\"FABRIKAM, INC.\"", StringComparison.Ordinal);
                if (hasPublisherCriteria)
                {
                    score += 1m;
                    result.Details.Add("Công thức có tiêu chí 'Fabrikam, Inc.'.");
                }
                else
                {
                    result.Errors.Add("Công thức F37 chưa dùng đúng tiêu chí publisher 'Fabrikam, Inc.'.");
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

// minor-sync: non-functional graders update
