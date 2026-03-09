using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T2Grader : ITaskGrader
    {
        public string TaskId => "P05-T2";
        public string TaskName => "Cot Difference tinh Selling Price - Cost, giu nguyen dinh dang";
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
                        result.Errors.Add($"Hang {row}: o H{row} chua co cong thuc.");
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
                            $"Hang {row}: cong thuc Difference phai la Selling Price - Cost (F{row}-G{row}). Hien tai: '{formula}'.");
                    }

                    var styleOk = diffCell.StyleID == ws.Cells[row, costCol].StyleID;
                    if (styleOk)
                    {
                        stylePreservedCount++;
                    }
                    else
                    {
                        result.Errors.Add(
                            $"Hang {row}: dinh dang cot Difference da thay doi (StyleID H{row}={diffCell.StyleID}, G{row}={ws.Cells[row, costCol].StyleID}).");
                    }
                }

                decimal score = 0;
                score += 0.5m; // Tim dung cot va range can cham.

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
                    result.Details.Add("Da dien cong thuc cho toan bo H5:H35.");
                }
                if (exactFormulaCount == expectedRows)
                {
                    result.Details.Add("Cong thuc Difference dung dang Selling Price - Cost cho tat ca dong.");
                }
                if (stylePreservedCount == expectedRows && headerStylePreserved)
                {
                    result.Details.Add("Dinh dang cot Difference duoc giu nguyen.");
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
