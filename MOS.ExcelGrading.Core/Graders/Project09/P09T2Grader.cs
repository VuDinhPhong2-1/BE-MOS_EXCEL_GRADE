using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project09
{
    public class P09T2Grader : ITaskGrader
    {
        public string TaskId => "P09-T2";
        public string TaskName => "Unmerge A1, apply Title style, 24pt, bold";
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
                var cell = studentSheet.Cells["A1"];
                decimal score = 0;

                // Rule 1: Không merge (1 điểm)
                if (!cell.Merge)
                {
                    score += 1;
                    result.Details.Add("✓ Đã hủy merge cell A1");
                }
                else
                {
                    result.Errors.Add("❌ Cell A1 vẫn còn merge");
                }

                // Rule 2: Style = Title (1 điểm)
                if (cell.StyleName.Contains("Title", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1;
                    result.Details.Add("✓ Áp dụng Title style");
                }
                else
                {
                    result.Errors.Add($"❌ Style không đúng (hiện tại: {cell.StyleName})");
                }

                // Rule 3: Font size = 24 (1 điểm)
                if (cell.Style.Font.Size == 24)
                {
                    score += 1;
                    result.Details.Add("✓ Font size 24pt");
                }
                else
                {
                    result.Errors.Add($"❌ Font size sai (hiện tại: {cell.Style.Font.Size}pt)");
                }

                // Rule 4: Bold (1 điểm)
                if (cell.Style.Font.Bold)
                {
                    score += 1;
                    result.Details.Add("✓ In đậm");
                }
                else
                {
                    result.Errors.Add("❌ Chưa in đậm");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"❌ Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}
