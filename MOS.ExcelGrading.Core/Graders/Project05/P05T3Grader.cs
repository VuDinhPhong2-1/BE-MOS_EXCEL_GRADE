using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T3Grader : ITaskGrader
    {
        public string TaskId => "P05-T3";
        public string TaskName => "F37 tinh trung binh Selling Price cho 'Fabrikam, Inc.'";
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
                    result.Errors.Add("Khong tim thay sheet 'Annual Purchases'.");
                    return result;
                }

                var cell = ws.Cells["F37"];
                var formulaRaw = cell.Formula;
                var formula = P05GraderHelpers.NormalizeFormula(formulaRaw);

                decimal score = 0;
                if (!string.IsNullOrWhiteSpace(formulaRaw))
                {
                    score += 1m;
                    result.Details.Add("O F37 da co cong thuc.");
                }
                else
                {
                    result.Errors.Add("O F37 chua co cong thuc.");
                    result.Score = score;
                    return result;
                }

                var usesAverage = formula.StartsWith("AVERAGEIF(", StringComparison.Ordinal) ||
                                  formula.StartsWith("AVERAGEIFS(", StringComparison.Ordinal);
                if (usesAverage)
                {
                    score += 1m;
                    result.Details.Add("Cong thuc da dung ham AVERAGEIF/AVERAGEIFS.");
                }
                else
                {
                    result.Errors.Add($"Cong thuc F37 chua dung ham AVERAGEIF. Hien tai: '{formulaRaw}'.");
                }

                var hasPublisherRange = formula.Contains("D5:D35", StringComparison.Ordinal);
                var hasSellingRange = formula.Contains("F5:F35", StringComparison.Ordinal);
                if (hasPublisherRange && hasSellingRange)
                {
                    score += 1m;
                    result.Details.Add("Cong thuc tham chieu dung cac vung D5:D35 va F5:F35.");
                }
                else
                {
                    result.Errors.Add(
                        $"Cong thuc F37 chua tham chieu dung range (can co D5:D35 va F5:F35). Formula: '{formulaRaw}'.");
                }

                var hasPublisherCriteria = formula.Contains("\"FABRIKAM,INC.\"", StringComparison.Ordinal) ||
                                           formula.Contains("\"FABRIKAM, INC.\"", StringComparison.Ordinal);
                if (hasPublisherCriteria)
                {
                    score += 1m;
                    result.Details.Add("Cong thuc co tieu chi 'Fabrikam, Inc.'.");
                }
                else
                {
                    result.Errors.Add("Cong thuc F37 chua dung dung tieu chi publisher 'Fabrikam, Inc.'.");
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
