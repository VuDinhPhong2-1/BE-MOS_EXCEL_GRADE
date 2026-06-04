using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T3Grader : ITaskGrader
    {
        public string TaskId => "P03-T3";
        public string TaskName => "Thêm Header bên phải 'Sequential' và để ở chế độ Normal";
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

                var rightHeaderRaw = ws.HeaderFooter?.OddHeader?.RightAlignedText ?? string.Empty;
                var rightText = P03GraderHelpers.ExtractRightHeaderText(rightHeaderRaw);
                if (string.Equals(rightText, "Sequential", StringComparison.Ordinal))
                {
                    result.Score += 3m;
                    result.Details.Add("Header bên phải đúng nội dung 'Sequential'");
                }
                else
                {
                    result.Errors.Add($"Header bên phải chưa đúng. Hiện tại: '{rightText}'");
                }

                var isNormalView = !ws.View.PageLayoutView && !ws.View.PageBreakView;
                if (isNormalView)
                {
                    result.Score += 1m;
                    result.Details.Add("Sheet đang ở chế độ Normal");
                }
                else
                {
                    result.Errors.Add("Sheet chưa ở chế độ Normal");
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
