using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T1Grader : ITaskGrader
    {
        public string TaskId => "P03-T1";
        public string TaskName => "Gop o A1:N1 tren sheet Ingredients";
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

                const string expectedTitle = "MUNSON'S PICKLES AND PRESERVES FARM";
                var a1Value = ws.Cells["A1"].Text?.Trim() ?? string.Empty;
                if (string.Equals(a1Value, expectedTitle, StringComparison.OrdinalIgnoreCase))
                {
                    result.Score += 1m;
                    result.Details.Add("O A1 van giu noi dung tieu de");
                }
                else
                {
                    result.Errors.Add($"Noi dung A1 da thay doi. Hien tai: '{a1Value}'");
                }

                var hasMergedA1N1 = ws.MergedCells.Any(r =>
                    string.Equals(r, "A1:N1", StringComparison.OrdinalIgnoreCase));

                if (hasMergedA1N1 && ws.Cells["A1"].Merge)
                {
                    result.Score += 2m;
                    result.Details.Add("Da gop dung vung A1:N1");
                }
                else
                {
                    result.Errors.Add("Chua gop dung vung A1:N1");
                }

                var alignment = ws.Cells["A1"].Style.HorizontalAlignment;
                var isCentered = alignment == ExcelHorizontalAlignment.Center
                                 || alignment == ExcelHorizontalAlignment.CenterContinuous;
                if (!isCentered)
                {
                    result.Score += 1m;
                    result.Details.Add("Canh ngang A1 khong bi doi sang Center (dung kieu Merge Across)");
                }
                else
                {
                    result.Errors.Add("A1 dang canh giua (co dau hieu dung Merge & Center thay vi Merge Across)");
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
