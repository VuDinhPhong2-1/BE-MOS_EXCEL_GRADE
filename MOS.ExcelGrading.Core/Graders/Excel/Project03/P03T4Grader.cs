using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T4Grader : ITaskGrader
    {
        public string TaskId => "P03-T4";
        public string TaskName => "Tạo hyperlink tại A6 đến Description!A18";
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
                    result.Errors.Add("Không tìm thấy sheet Ingredients");
                    return result;
                }

                var cell = ws.Cells["A6"];
                var hyperlink = cell.Hyperlink;
                if (hyperlink == null)
                {
                    result.Errors.Add("Ô A6 chưa có hyperlink");
                    return result;
                }

                result.Score += 1m;
                result.Details.Add("A6 đã có hyperlink");

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
                    result.Details.Add("Đích đến hyperlink đúng Description!A18");
                }
                else
                {
                    result.Errors.Add("Đích đến hyperlink chưa đúng Description!A18");
                }

                var displayOk = !string.IsNullOrWhiteSpace(cell.Text);
                if (displayOk)
                {
                    result.Score += 1m;
                    result.Details.Add("Nội dung hiển thị tại A6 được giữ lại");
                }
                else
                {
                    result.Errors.Add("Nội dung hiển thị tại A6 bị trống");
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
