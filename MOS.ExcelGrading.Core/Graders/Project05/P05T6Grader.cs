using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T6Grader : ITaskGrader
    {
        public string TaskId => "P05-T6";
        public string TaskName => "Mở cửa sổ thứ hai và hiển thị Side by Side dạng trên-dưới";
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
                var views = P05GraderHelpers.GetWorkbookViewNodes(studentSheet);
                if (views == null || views.Count == 0)
                {
                    result.Errors.Add("Không đọc được thông tin workbookView.");
                    return result;
                }

                decimal score = 0;
                if (views.Count >= 2)
                {
                    score += 2m;
                    result.Details.Add($"Workbook có {views.Count} cửa sổ view.");
                }
                else
                {
                    result.Errors.Add("Workbook chỉ có 1 cửa sổ view (chưa mở cửa sổ thứ hai).");
                    result.Score = score;
                    return result;
                }

                var first = views[0];
                var second = views[1];

                var hasX1 = P05GraderHelpers.TryParseDoubleAttribute(first, "xWindow", out var x1);
                var hasX2 = P05GraderHelpers.TryParseDoubleAttribute(second, "xWindow", out var x2);
                var hasW1 = P05GraderHelpers.TryParseDoubleAttribute(first, "windowWidth", out var w1);
                var hasW2 = P05GraderHelpers.TryParseDoubleAttribute(second, "windowWidth", out var w2);
                var sameLeft = hasX1 && hasX2 && Math.Abs(x1 - x2) <= 30;
                var sameWidth = hasW1 && hasW2 && Math.Abs(w1 - w2) <= 120;

                if (sameLeft && sameWidth)
                {
                    score += 1m;
                    result.Details.Add("Hai cửa sổ có cùng vị trí ngang và độ rộng tương đương.");
                }
                else
                {
                    result.Errors.Add("Hai cửa sổ chưa canh theo chiều trên-dưới (xWindow/windowWidth chưa đồng nhất).");
                }

                var hasY1 = P05GraderHelpers.TryParseDoubleAttribute(first, "yWindow", out var y1);
                var hasY2 = P05GraderHelpers.TryParseDoubleAttribute(second, "yWindow", out var y2);
                if (hasY1 && hasY2 && Math.Abs(y1 - y2) > 50 && y2 > y1)
                {
                    score += 1m;
                    result.Details.Add("Cửa sổ thứ hai nằm bên dưới cửa sổ thứ nhất (top-bottom side by side).");
                }
                else
                {
                    result.Errors.Add("Không thấy dấu hiệu sắp xếp trên-dưới cho 2 cửa sổ workbook.");
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
