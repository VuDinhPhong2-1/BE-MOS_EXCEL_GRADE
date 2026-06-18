using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project05
{
    public class P05T4Grader : ITaskGrader
    {
        public string TaskId => "P05-T4";
        public string TaskName => "Đặt sheet 'Annual Purchases' vào giữa 'Works' và 'Titles'";
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
                var names = studentSheet.Workbook.Worksheets
                    .Select(w => (w.Name ?? string.Empty).Trim())
                    .ToList();

                var worksIndex = names.FindIndex(n => string.Equals(n, "Works", StringComparison.OrdinalIgnoreCase));
                var annualIndex = names.FindIndex(n => string.Equals(n, "Annual Purchases", StringComparison.OrdinalIgnoreCase));
                var titlesIndex = names.FindIndex(n => string.Equals(n, "Titles", StringComparison.OrdinalIgnoreCase));

                decimal score = 0;
                if (worksIndex >= 0 && annualIndex >= 0 && titlesIndex >= 0)
                {
                    score += 1m;
                    result.Details.Add("Đã tồn tại đầy đủ 3 sheet Works, Annual Purchases, Titles.");
                }
                else
                {
                    result.Errors.Add("Thiếu một trong các sheet bắt buộc: Works / Annual Purchases / Titles.");
                    result.Score = score;
                    return result;
                }

                if (worksIndex < annualIndex && annualIndex < titlesIndex)
                {
                    score += 1m;
                    result.Details.Add("Thứ tự tổng quát đúng: Works -> Annual Purchases -> Titles.");
                }
                else
                {
                    result.Errors.Add("Thứ tự tổng quát sheet chưa đúng.");
                }

                if (annualIndex == worksIndex + 1)
                {
                    score += 1m;
                    result.Details.Add("Annual Purchases đứng ngay sau Works.");
                }
                else
                {
                    result.Errors.Add("Annual Purchases chưa nằm ngay sau Works.");
                }

                if (titlesIndex == annualIndex + 1)
                {
                    score += 1m;
                    result.Details.Add("Titles đứng ngay sau Annual Purchases.");
                }
                else
                {
                    result.Errors.Add("Titles chưa nằm ngay sau Annual Purchases.");
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

// minor-sync: non-functional graders update

