using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project11
{
    public class P11T1Grader : ITaskGrader
    {
        public string TaskId => "P11-T1";
        public string TaskName => "Games: Merge A12:B12 den A18:B18";
        public decimal MaxScore => 3;

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
                var ws = P11GraderHelpers.GetSheet(studentSheet.Workbook, "Games");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Games'.");
                    return result;
                }

                decimal score = 0m;
                var expectedRanges = Enumerable.Range(12, 7)
                    .Select(row => $"A{row}:B{row}")
                    .ToList();

                var matchedCount = expectedRanges.Count(range => P11GraderHelpers.HasMergeRange(ws, range));
                if (matchedCount == expectedRanges.Count)
                {
                    score += 2m;
                    result.Details.Add("Ðã merge d?y d? các vùng A12:B12 den A18:B18.");
                }
                else
                {
                    var missing = expectedRanges.Where(range => !P11GraderHelpers.HasMergeRange(ws, range));
                    result.Errors.Add($"Thi?u merge ranges: {string.Join(", ", missing)}.");
                }

                var centeredCount = expectedRanges.Count(range =>
                {
                    var topLeft = range.Split(':')[0];
                    return P11GraderHelpers.IsMergedCellCentered(ws, topLeft);
                });
                if (centeredCount == expectedRanges.Count)
                {
                    score += 1m;
                    result.Details.Add("Các ô merge dã can gi?a dúng.");
                }

                var hasOnlyExpectedMerges = ws.MergedCells.Count == expectedRanges.Count;
                if (hasOnlyExpectedMerges)
                {
                    score += 1m;
                    result.Details.Add("S? lu?ng merge range dung (7).");
                }
                else
                {
                    result.Errors.Add($"S? lu?ng merge range chua dúng. Hi?n t?i: {ws.MergedCells.Count}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}


