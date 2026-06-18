using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project03
{
    public class P03T2Grader : ITaskGrader
    {
        public string TaskId => "P03-T2";
        public string TaskName => "AutoFit tất cả cột A:N trên sheet Ingredients";
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
                var ws = P03GraderHelpers.GetIngredientsSheet(studentSheet);
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet Ingredients");
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
                    result.Details.Add("Tất cả cột A:N có BestFit=true");
                }
                else
                {
                    var partial = Math.Round(3m * bestFitCount / totalCols, 2);
                    result.Score += partial;
                    result.Errors.Add($"BestFit chưa đầy đủ ({bestFitCount}/{totalCols})");
                }

                if (nonDefaultWidthCount == totalCols)
                {
                    result.Score += 1m;
                    result.Details.Add("Tất cả cột A:N đã thay đổi độ rộng (không còn mặc định 9)");
                }
                else
                {
                    result.Errors.Add($"Số cột có độ rộng khác mặc định: {nonDefaultWidthCount}/{totalCols}");
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


// minor-sync: non-functional graders update

