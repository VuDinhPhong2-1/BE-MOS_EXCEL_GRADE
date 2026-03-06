using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T4Grader : ITaskGrader
    {
        public string TaskId => "P03-T4";
        public string TaskName => "Tao hyperlink tai A6 den Description!A18";
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

                var cell = ws.Cells["A6"];
                var hyperlink = cell.Hyperlink;
                if (hyperlink == null)
                {
                    result.Errors.Add("O A6 chua co hyperlink");
                    return result;
                }

                result.Score += 1m;
                result.Details.Add("A6 da co hyperlink");

                var targetOk = false;
                if (hyperlink is ExcelHyperLink excelLink)
                {
                    var normalized = P03GraderHelpers.NormalizeRef(excelLink.ReferenceAddress);
                    targetOk = normalized == "DESCRIPTION!A18";
                }
                else
                {
                    var normalized = P03GraderHelpers.NormalizeRef(hyperlink.OriginalString);
                    targetOk = normalized.Contains("DESCRIPTION!A18", StringComparison.Ordinal);
                }

                if (targetOk)
                {
                    result.Score += 2m;
                    result.Details.Add("Dich den hyperlink dung Description!A18");
                }
                else
                {
                    result.Errors.Add("Dich den hyperlink chua dung Description!A18");
                }

                var displayOk = !string.IsNullOrWhiteSpace(cell.Text);
                if (displayOk)
                {
                    result.Score += 1m;
                    result.Details.Add("Noi dung hien thi tai A6 duoc giu lai");
                }
                else
                {
                    result.Errors.Add("Noi dung hien thi tai A6 bi trong");
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

