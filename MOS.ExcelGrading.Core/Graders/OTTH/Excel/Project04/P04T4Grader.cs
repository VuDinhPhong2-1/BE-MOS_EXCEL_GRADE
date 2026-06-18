using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project04
{
    public class P04T4Grader : ITaskGrader
    {
        public string TaskId => "P04-T4";
        public string TaskName => "Convert table trên Classes thành range (giữ định dạng)";
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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Classes");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet Classes");
                    return result;
                }

                if (ws.Tables.Count == 0)
                {
                    result.Score += 2m;
                    result.Details.Add("Đã convert table thành range (không còn table object)");
                }
                else
                {
                    result.Errors.Add($"Vẫn còn {ws.Tables.Count} table trên sheet Classes");
                }

                var expectedHeaders = new[] { "Title", "Section", "Days", "Hours", "Location", "Instructor" };
                var headersOk = true;
                for (var i = 0; i < expectedHeaders.Length; i++)
                {
                    var actual = ws.Cells[4, i + 1].Text.Trim();
                    if (!string.Equals(actual, expectedHeaders[i], StringComparison.OrdinalIgnoreCase))
                    {
                        headersOk = false;
                        break;
                    }
                }

                if (headersOk)
                {
                    result.Score += 1m;
                    result.Details.Add("Header A4:F4 vẫn đầy đủ sau khi convert");
                }
                else
                {
                    result.Errors.Add("Header A4:F4 chưa đúng sau khi convert");
                }

                var headerStyleOk = ws.Cells[4, 1].StyleID > 0
                                    && ws.Cells[4, 2].StyleID > 0
                                    && ws.Cells[4, 6].StyleID > 0;
                var dataStyleOk = ws.Cells[5, 1].StyleID > 0
                                  && ws.Cells[5, 2].StyleID > 0
                                  && ws.Cells[25, 6].StyleID > 0;
                if (headerStyleOk && dataStyleOk)
                {
                    result.Score += 1m;
                    result.Details.Add("Định dạng cơ bản được giữ lại sau khi convert");
                }
                else
                {
                    result.Errors.Add("Định dạng cơ bản có dấu hiệu bị mất sau khi convert");
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

