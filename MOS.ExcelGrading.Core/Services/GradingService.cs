using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Graders.Project09;

namespace MOS.ExcelGrading.Core.Services
{
    public class GradingService : IGradingService
    {
        private readonly List<ITaskGrader> _project09Graders;

        public GradingService()
        {
            _project09Graders = new List<ITaskGrader>
            {
                new P09T1Grader(),
                new P09T2Grader(),
                new P09T3Grader(),
                new P09T4Grader(),
                new P09T5Grader(),
                new P09T6Grader()
            };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<GradingResult> GradeProject09Async(Stream studentFile, Stream answerFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P09",
                ProjectName = "Sales and Orders Report"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                using var answerPackage = new ExcelPackage(answerFile);

                var studentSheet = studentPackage.Workbook.Worksheets[0];
                var answerSheet = answerPackage.Workbook.Worksheets[0];

                foreach (var grader in _project09Graders)
                {
                    var taskResult = await Task.Run(() =>
                        grader.Grade(studentSheet, answerSheet)
                    );
                    result.TaskResults.Add(taskResult);
                }

                result.TotalScore = result.TaskResults.Sum(t => t.Score);
                result.MaxScore = result.TaskResults.Sum(t => t.MaxScore);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "System Error",
                    Errors = new List<string> { $"❌ {ex.Message}" }
                });
            }

            return result;
        }
    }
}
