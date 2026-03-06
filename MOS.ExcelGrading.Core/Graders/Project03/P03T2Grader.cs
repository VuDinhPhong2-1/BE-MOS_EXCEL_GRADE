using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T2Grader : ITaskGrader
    {
        public string TaskId => "P03-T2";
        public string TaskName => "AutoFit tat ca cot A:N tren Ingredients";
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
                var ws = P03GraderHelpers.GetIngredientsSheet(studentSheet);
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet Ingredients");
                    return result;
                }

                const int firstCol = 1;  // A
                const int lastCol = 14;  // N
                const int totalCols = lastCol - firstCol + 1;

                var bestFitCount = 0;
                var nonDefaultWidthCount = 0;

                for (var c = firstCol; c <= lastCol; c++)
                {
                    var col = ws.Column(c);
                    if (col.BestFit)
                    {
                        bestFitCount++;
                    }

                    if (Math.Abs(col.Width - 9d) > 0.01d)
                    {
                        nonDefaultWidthCount++;
                    }
                }

                if (bestFitCount == totalCols)
                {
                    result.Score += 3m;
                    result.Details.Add("Tat ca cot A:N co BestFit=true");
                }
                else
                {
                    var partial = Math.Round(3m * bestFitCount / totalCols, 2);
                    result.Score += partial;
                    result.Errors.Add($"BestFit chua day du ({bestFitCount}/{totalCols})");
                }

                if (nonDefaultWidthCount == totalCols)
                {
                    result.Score += 1m;
                    result.Details.Add("Tat ca cot A:N da thay doi do rong (khong con mac dinh 9)");
                }
                else
                {
                    result.Errors.Add($"So cot co do rong khac mac dinh: {nonDefaultWidthCount}/{totalCols}");
                }

                result.Score = Math.Min(MaxScore, result.Score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}

