using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project05
{
    public class P05T1Grader : ITaskGrader
    {
        public string TaskId => "P05-T1";
        public string TaskName => "Đặt Document Properties Company = 'Salon International'";
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
                var company = studentSheet.Workbook.Properties.Company ?? string.Empty;
                var expected = "Salon International";

                if (string.Equals(company, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Company đã được đặt chính xác là 'Salon International'.");
                }
                else
                {
                    result.Errors.Add($"Company chưa đúng chính xác. Hiện tại: '{company}'. Mong đợi đúng chính tả: '{expected}'.");
                }
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

