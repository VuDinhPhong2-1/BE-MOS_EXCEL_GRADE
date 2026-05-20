using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project13
{
    public class P13T5Grader : ITaskGrader
    {
        public string TaskId => "P13-T5";
        public string TaskName => "Attendees: che do Page Layout View + ngat trang";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Attendees");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Attendees'.");
                    return result;
                }

                decimal score = 0m;
                if (ws.View.PageLayoutView)
                {
                    score += 2m;
                    result.Details.Add("Sheet dang o che do Page Layout.");
                }
                else
                {
                    result.Errors.Add("Sheet chưa ở che do Page Layout.");
                }

                var actualBreakIds = P13GraderHelpers.GetDetectedRowBreakIds(ws);
                const int expectedAnchorRow = 35;
                const int tolerance = 1;
                var hasExpectedBreak = actualBreakIds.Any(actual => Math.Abs(actual - expectedAnchorRow) <= tolerance);
                var hasOnlyAcceptedBreaks = actualBreakIds.All(actual => Math.Abs(actual - expectedAnchorRow) <= tolerance);

                if (hasExpectedBreak && hasOnlyAcceptedBreaks)
                {
                    score += 2m;
                    result.Details.Add(
                        $"Da chen page break dung vi tri quanh dòng {expectedAnchorRow} (actual id={P13GraderHelpers.FormatIdList(actualBreakIds)}).");
                }
                else
                {
                    result.Errors.Add(
                        $"Page break chưa đúng. Mong doi co break quanh dòng {expectedAnchorRow} (chap nhan id {expectedAnchorRow - tolerance}..{expectedAnchorRow + tolerance}) va khong co break khac. Actual id={P13GraderHelpers.FormatIdList(actualBreakIds)}.");
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



