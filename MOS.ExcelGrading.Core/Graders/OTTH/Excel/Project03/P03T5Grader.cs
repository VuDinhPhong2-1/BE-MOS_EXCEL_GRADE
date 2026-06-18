using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project03
{
    public class P03T5Grader : ITaskGrader
    {
        public string TaskId => "P03-T5";
        public string TaskName => "Print settings: Landscape + fit all columns on one page";
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

                var ps = ws.PrinterSettings;
                var isLandscape = ps.Orientation == eOrientation.Landscape;
                if (isLandscape)
                {
                    result.Score += 1m;
                    result.Details.Add("Đã đặt hướng in Landscape");
                }
                else
                {
                    result.Errors.Add("Chưa đặt hướng in Landscape");
                }

                var fitToPage = ps.FitToPage;
                if (fitToPage)
                {
                    result.Score += 1.5m;
                    result.Details.Add("Đã bật FitToPage");
                }
                else
                {
                    result.Errors.Add("Chưa bật FitToPage");
                }

                var fitToHeightZero = ps.FitToHeight == 0;
                var fitToWidthOneOrDefault = ps.FitToWidth == 1 || ps.FitToWidth == int.MinValue;
                if (fitToHeightZero && fitToWidthOneOrDefault)
                {
                    result.Score += 1.5m;
                    result.Details.Add("Cấu hình in 1 trang theo chiều ngang (fit columns)");
                }
                else
                {
                    result.Errors.Add($"Thông số fit chưa đúng (FitToWidth={ps.FitToWidth}, FitToHeight={ps.FitToHeight})");
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

