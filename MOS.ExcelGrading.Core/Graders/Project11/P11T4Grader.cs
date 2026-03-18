using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T4Grader : ITaskGrader
    {
        public string TaskId => "P11-T4";
        public string TaskName => "Games: keep only table A2:F9";
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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Games");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Games'.");
                    return result;
                }

                decimal score = 0m;
                var topTable = P11GraderHelpers.FindTableByAddress(ws, "A2:F9");
                if (topTable != null)
                {
                    score += 2m;
                    result.Details.Add("Da giu table A2:F9.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay table A2:F9. Hien tai: {P11GraderHelpers.JoinTableAddresses(ws)}");
                }

                if (ws.Tables.Count == 1)
                {
                    score += 2m;
                    result.Details.Add("Chi con 1 table tren sheet Games.");
                }
                else
                {
                    result.Errors.Add($"So luong table chua dung. Hien tai: {ws.Tables.Count}.");
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
