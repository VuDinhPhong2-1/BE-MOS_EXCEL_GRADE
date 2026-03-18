using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project11
{
    public class P11T3Grader : ITaskGrader
    {
        public string TaskId => "P11-T3";
        public string TaskName => "Games: merge A12:B12 through A18:B18";
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
                var expected = Enumerable.Range(12, 7).Select(row => $"A{row}:B{row}").ToHashSet(StringComparer.OrdinalIgnoreCase);
                var actual = ws.MergedCells
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                if (expected.All(actual.Contains))
                {
                    score += 2m;
                    result.Details.Add("Da merge day du cac vung A12:B12 den A18:B18.");
                }
                else
                {
                    var missing = expected.Where(x => !actual.Contains(x));
                    result.Errors.Add($"Thieu merge ranges: {string.Join(", ", missing)}.");
                }

                if (actual.Count == expected.Count)
                {
                    score += 1m;
                    result.Details.Add("So luong merge ranges dung (7).");
                }
                else
                {
                    result.Errors.Add($"So luong merge ranges chua dung. Hien tai: {actual.Count}.");
                }

                var allTwoColumnMerge = actual.All(range =>
                {
                    var parts = range.Split(':');
                    return parts.Length == 2
                           && parts[0].StartsWith("A", StringComparison.OrdinalIgnoreCase)
                           && parts[1].StartsWith("B", StringComparison.OrdinalIgnoreCase);
                });
                if (allTwoColumnMerge)
                {
                    score += 1m;
                    result.Details.Add("Merge format dung dang 2 cot A:B.");
                }
                else
                {
                    result.Errors.Add("Co merge range khong nam trong dang A#:B#.");
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
