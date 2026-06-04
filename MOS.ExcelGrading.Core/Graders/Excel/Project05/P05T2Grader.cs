using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T2Grader : ITaskGrader
    {
        public string TaskId => "P05-T2";
        public string TaskName => "Cột Difference tính Selling Price - Cost, giữ nguyên định dạng";
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
                var ws = P05GraderHelpers.GetSheet(studentSheet, "Annual Purchases");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Annual Purchases'.");
                    return result;
                }

                const int costCol = 7;         // G
                const int diffCol = 8;         // H
                const int dataStartRow = 5;
                const int dataEndRow = 35;

                var formulaCount = 0;
                var exactFormulaCount = 0;
                var stylePreservedCount = 0;

                var expectedRows = dataEndRow - dataStartRow + 1;
                var headerStylePreserved = ws.Cells[4, diffCol].StyleID == ws.Cells[4, costCol].StyleID;

                for (var row = dataStartRow; row <= dataEndRow; row++)
                {
                    var diffCell = ws.Cells[row, diffCol];
                    var formula = diffCell.Formula;
                    var formulaR1C1 = diffCell.FormulaR1C1;

                    var hasFormula =
                        !string.IsNullOrWhiteSpace(formula) ||
                        !string.IsNullOrWhiteSpace(formulaR1C1);
                    if (hasFormula)
                    {
                        formulaCount++;
                    }
                    else
                    {
                        result.Errors.Add($"Hàng {row}: o H{row} chưa có công thức.");
                    }

                    var formulaOk =
                        P05GraderHelpers.IsDifferenceFormula(formula, row) ||
                        P05GraderHelpers.IsDifferenceFormulaR1C1(formulaR1C1);
                    if (formulaOk)
                    {
                        exactFormulaCount++;
                    }
                    else if (hasFormula)
                    {
                        result.Errors.Add(
                            $"Hàng {row}: công thức Difference phải là Selling Price - Cost (F{row}-G{row}). Hiện tại: '{formula}'.");
                    }

                    var styleOk = diffCell.StyleID == ws.Cells[row, costCol].StyleID;
                    if (styleOk)
                    {
                        stylePreservedCount++;
                    }
                    else
                    {
                        result.Errors.Add(
                            $"Hàng {row}: định dạng cột Difference đã thay đổi (StyleID H{row}={diffCell.StyleID}, G{row}={ws.Cells[row, costCol].StyleID}).");
                    }
                }

                decimal score = 0;
                score += 0.5m; // Tìm đúng cột và range cần chấm.

                score += Math.Round(1.5m * formulaCount / expectedRows, 2);
                score += Math.Round(1.5m * exactFormulaCount / expectedRows, 2);

                var styleComponent = Math.Round(0.5m * stylePreservedCount / expectedRows, 2);
                if (headerStylePreserved)
                {
                    styleComponent = Math.Min(0.5m, styleComponent + 0.05m);
                }
                score += Math.Min(0.5m, styleComponent);

                if (formulaCount == expectedRows)
                {
                    result.Details.Add("Đã điền công thức cho toàn bộ H5:H35.");
                }
                if (exactFormulaCount == expectedRows)
                {
                    result.Details.Add("Công thức Difference đúng dạng Selling Price - Cost cho tất cả dòng.");
                }
                if (stylePreservedCount == expectedRows && headerStylePreserved)
                {
                    result.Details.Add("Định dạng cột Difference được giữ nguyên.");
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
