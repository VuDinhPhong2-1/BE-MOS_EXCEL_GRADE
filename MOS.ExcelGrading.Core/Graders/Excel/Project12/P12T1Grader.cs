using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace MOS.ExcelGrading.Core.Graders.Project12
{
    public class P12T1Grader : ITaskGrader
    {
        public string TaskId => "P12-T1";
        public string TaskName => "Range: Merge E7:F7";
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
                var ws = P12GraderHelpers.GetSheet(studentSheet.Workbook, "Range");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Range'.");
                    return result;
                }

                decimal score = 0m;
                if (P12GraderHelpers.HasMergeRange(ws, "E7:F7"))
                {
                    score += 2m;
                    result.Details.Add("Da merge dung vung E7:F7.");
                }
                else
                {
                    result.Errors.Add("Chua merge dung vung E7:F7.");
                }

                if (!P12GraderHelpers.HasMergeRange(ws, "E8:F8"))
                {
                    score += 1m;
                    result.Details.Add("Không còn merge cu E8:F8.");
                }
                else
                {
                    result.Errors.Add("Van con merge cu E8:F8.");
                }

                var styleMoved = ws.Cells["E7"].StyleID > 0 && ws.Cells["F7"].StyleID > 0;
                if (styleMoved)
                {
                    score += 1m;
                    result.Details.Add("Định dạng cua o merge duoc giu lai.");
                }
                else
                {
                    result.Errors.Add("Định dạng o merge E7:F7 khong hop le.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}



