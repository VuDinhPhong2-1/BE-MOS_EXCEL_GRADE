using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T3Grader : ITaskGrader
    {
        public string TaskId => "P03-T3";
        public string TaskName => "Them Header ben phai 'Sequential' va de o che do Normal";
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

                var rightHeaderRaw = ws.HeaderFooter?.OddHeader?.RightAlignedText ?? string.Empty;
                var rightText = P03GraderHelpers.ExtractRightHeaderText(rightHeaderRaw);
                if (string.Equals(rightText, "Sequential", StringComparison.OrdinalIgnoreCase))
                {
                    result.Score += 3m;
                    result.Details.Add("Header ben phai dung noi dung 'Sequential'");
                }
                else
                {
                    result.Errors.Add($"Header ben phai chua dung. Hien tai: '{rightText}'");
                }

                var isNormalView = !ws.View.PageLayoutView && !ws.View.PageBreakView;
                if (isNormalView)
                {
                    result.Score += 1m;
                    result.Details.Add("Sheet dang o che do Normal");
                }
                else
                {
                    result.Errors.Add("Sheet chua o che do Normal");
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

