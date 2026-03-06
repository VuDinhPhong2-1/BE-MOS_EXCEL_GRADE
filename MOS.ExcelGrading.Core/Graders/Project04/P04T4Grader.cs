using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T4Grader : ITaskGrader
    {
        public string TaskId => "P04-T4";
        public string TaskName => "Convert table tren Classes thanh range (giu dinh dang)";
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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Classes");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet Classes");
                    return result;
                }

                if (ws.Tables.Count == 0)
                {
                    result.Score += 2m;
                    result.Details.Add("Da convert table thanh range (khong con table object)");
                }
                else
                {
                    result.Errors.Add($"Van con {ws.Tables.Count} table tren sheet Classes");
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
                    result.Details.Add("Header A4:F4 van day du sau khi convert");
                }
                else
                {
                    result.Errors.Add("Header A4:F4 chua dung sau khi convert");
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
                    result.Details.Add("Dinh dang co ban duoc giu lai sau khi convert");
                }
                else
                {
                    result.Errors.Add("Dinh dang co ban co dau hieu bi mat sau khi convert");
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

