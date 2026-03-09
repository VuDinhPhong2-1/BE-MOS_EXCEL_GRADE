using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T4Grader : ITaskGrader
    {
        public string TaskId => "P05-T4";
        public string TaskName => "Dat sheet 'Annual Purchases' vao giua 'Works' va 'Titles'";
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
                    result.Details.Add("Da ton tai day du 3 sheet Works, Annual Purchases, Titles.");
                }
                else
                {
                    result.Errors.Add("Thieu mot trong cac sheet bat buoc: Works / Annual Purchases / Titles.");
                    result.Score = score;
                    return result;
                }

                if (worksIndex < annualIndex && annualIndex < titlesIndex)
                {
                    score += 1m;
                    result.Details.Add("Thu tu tong quat dung: Works -> Annual Purchases -> Titles.");
                }
                else
                {
                    result.Errors.Add("Thu tu tong quat sheet chua dung.");
                }

                if (annualIndex == worksIndex + 1)
                {
                    score += 1m;
                    result.Details.Add("Annual Purchases dung ngay sau Works.");
                }
                else
                {
                    result.Errors.Add("Annual Purchases chua nam ngay sau Works.");
                }

                if (titlesIndex == annualIndex + 1)
                {
                    score += 1m;
                    result.Details.Add("Titles dung ngay sau Annual Purchases.");
                }
                else
                {
                    result.Errors.Add("Titles chua nam ngay sau Annual Purchases.");
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
