using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T5Grader : ITaskGrader
    {
        public string TaskId => "P11-T5";
        public string TaskName => "Thiet lap ti le in cho tung worksheet (Fit to one page)";
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
                var worksheets = studentSheet.Workbook.Worksheets
                    .Where(ws => ws is not OfficeOpenXml.ExcelChartsheet)
                    .ToList();
                if (worksheets.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy worksheet de kiem tra print scaling.");
                    return result;
                }

                decimal score = 0m;
                var fitToPageCount = 0;
                foreach (var ws in worksheets)
                {
                    var ns = P11GraderHelpers.CreateWorkbookNamespaceManager(ws.WorksheetXml);
                    var pageSetUpPr = ws.WorksheetXml.SelectSingleNode("//x:sheetPr/x:pageSetUpPr", ns);
                    var fitToPage = pageSetUpPr?.Attributes?["fitToPage"]?.Value ?? string.Empty;
                    if (string.Equals(fitToPage, "1", StringComparison.Ordinal))
                    {
                        fitToPageCount++;
                    }
                }

                if (fitToPageCount == worksheets.Count)
                {
                    score += 3m;
                    result.Details.Add($"Da bat fitToPage cho tat ca worksheet ({fitToPageCount}/{worksheets.Count}).");
                }
                else
                {
                    result.Errors.Add($"Chua bat fitToPage day du ({fitToPageCount}/{worksheets.Count}).");
                }

                if (worksheets.Count == 4)
                {
                    score += 1m;
                    result.Details.Add("So worksheet dung 4 trang can cau hinh in.");
                }
                else
                {
                    result.Errors.Add($"So worksheet hien tai la {worksheets.Count}, mong đợi 4.");
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



