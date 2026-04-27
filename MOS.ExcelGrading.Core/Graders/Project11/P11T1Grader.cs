using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T1Grader : ITaskGrader
    {
        public string TaskId => "P11-T1";
        public string TaskName => "Games: Merge A12:B12 den A18:B18";
        public decimal MaxScore => 3;

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
                    result.Errors.Add("Không tìm thấy sheet 'Games'.");
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
                    result.Details.Add("Đã merge đầy đủ các vùng A12:B12 den A18:B18.");
                }
                else
                {
                    var missing = expectedRanges.Where(range => !P11GraderHelpers.HasMergeRange(ws, range));
                    result.Errors.Add($"Thiếu merge ranges: {string.Join(", ", missing)}.");
                }

                var centeredCount = expectedRanges.Count(range =>
                {
                    var topLeft = range.Split(':')[0];
                    return P11GraderHelpers.IsMergedCellCentered(ws, topLeft);
                });
                if (centeredCount == expectedRanges.Count)
                {
                    score += 1m;
                    result.Details.Add("Các ô merge đã căn giữa đúng.");
                }

                var hasOnlyExpectedMerges = ws.MergedCells.Count == expectedRanges.Count;
                if (hasOnlyExpectedMerges)
                {
                    score += 1m;
                    result.Details.Add("Số lượng merge range dung (7).");
                }
                else
                {
                    result.Errors.Add($"Số lượng merge range chưa đúng. Hiện tại: {ws.MergedCells.Count}.");
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

