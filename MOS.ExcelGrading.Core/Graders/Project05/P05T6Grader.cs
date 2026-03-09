using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T6Grader : ITaskGrader
    {
        public string TaskId => "P05-T6";
        public string TaskName => "Mo cua so thu hai va hien thi Side by Side dang tren-duoi";
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
                    result.Errors.Add("Khong doc duoc thong tin workbookView.");
                    return result;
                }

                decimal score = 0;
                if (views.Count >= 2)
                {
                    score += 2m;
                    result.Details.Add($"Workbook co {views.Count} cua so view.");
                }
                else
                {
                    result.Errors.Add("Workbook chi co 1 cua so view (chua mo cua so thu hai).");
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
                    result.Details.Add("Hai cua so co cung vi tri ngang va do rong tuong duong.");
                }
                else
                {
                    result.Errors.Add("Hai cua so chua canh theo chieu tren-duoi (xWindow/windowWidth chua dong nhat).");
                }

                var hasY1 = P05GraderHelpers.TryParseDoubleAttribute(first, "yWindow", out var y1);
                var hasY2 = P05GraderHelpers.TryParseDoubleAttribute(second, "yWindow", out var y2);
                if (hasY1 && hasY2 && Math.Abs(y1 - y2) > 50 && y2 > y1)
                {
                    score += 1m;
                    result.Details.Add("Cua so thu hai nam ben duoi cua so thu nhat (top-bottom side by side).");
                }
                else
                {
                    result.Errors.Add("Khong thay dau hieu sap xep tren-duoi cho 2 cua so workbook.");
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
