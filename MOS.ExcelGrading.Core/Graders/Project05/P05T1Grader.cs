using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project05
{
    public class P05T1Grader : ITaskGrader
    {
        public string TaskId => "P05-T1";
        public string TaskName => "Dat Document Properties Company = 'Salon International'";
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
                var company = studentSheet.Workbook.Properties.Company ?? string.Empty;
                var expected = "Salon International";

                if (string.Equals(company, expected, StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Company da duoc dat chinh xac la 'Salon International'.");
                }
                else
                {
                    result.Errors.Add($"Company chua dung chinh xac. Hien tai: '{company}'. Mong doi dung chinh ta: '{expected}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
