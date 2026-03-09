using OfficeOpenXml;
using MOS.ExcelGrading.Core.Graders.Project01;
using MOS.ExcelGrading.Core.Graders.Project02;
using MOS.ExcelGrading.Core.Graders.Project03;
using MOS.ExcelGrading.Core.Graders.Project04;
using MOS.ExcelGrading.Core.Graders.Project05;
using MOS.ExcelGrading.Core.Graders.Project06;
using MOS.ExcelGrading.Core.Graders.Project07;
using MOS.ExcelGrading.Core.Graders.Project08;
using MOS.ExcelGrading.Core.Graders.Project09;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class GradingService : IGradingService
    {
        private readonly List<ITaskGrader> _project01Graders;
        private readonly List<ITaskGrader> _project02Graders;
        private readonly List<ITaskGrader> _project03Graders;
        private readonly List<ITaskGrader> _project04Graders;
        private readonly List<ITaskGrader> _project05Graders;
        private readonly List<ITaskGrader> _project06Graders;
        private readonly List<ITaskGrader> _project07Graders;
        private readonly List<ITaskGrader> _project08Graders;
        private readonly List<ITaskGrader> _project09Graders;

        public GradingService()
        {
            _project01Graders = new List<ITaskGrader>
            {
                new P01T1Grader(),
                new P01T2Grader(),
                new P01T3Grader(),
                new P01T4Grader(),
                new P01T5Grader()
            };

            _project02Graders = new List<ITaskGrader>
            {
                new P02T1Grader(),
                new P02T2Grader(),
                new P02T3Grader(),
                new P02T4Grader(),
                new P02T5Grader(),
                new P02T6Grader(),
                new P02T7Grader()
            };

            _project03Graders = new List<ITaskGrader>
            {
                new P03T1Grader(),
                new P03T2Grader(),
                new P03T3Grader(),
                new P03T4Grader(),
                new P03T5Grader(),
                new P03T6Grader()
            };

            _project04Graders = new List<ITaskGrader>
            {
                new P04T1Grader(),
                new P04T2Grader(),
                new P04T3Grader(),
                new P04T4Grader(),
                new P04T5Grader(),
                new P04T6Grader(),
                new P04T7Grader()
            };

            _project05Graders = new List<ITaskGrader>
            {
                new P05T1Grader(),
                new P05T2Grader(),
                new P05T3Grader(),
                new P05T4Grader(),
                new P05T5Grader(),
                new P05T6Grader()
            };

            _project06Graders = new List<ITaskGrader>
            {
                new P06T1Grader(),
                new P06T2Grader(),
                new P06T3Grader(),
                new P06T4Grader(),
                new P06T5Grader(),
                new P06T6Grader()
            };

            _project07Graders = new List<ITaskGrader>
            {
                new P07T1Grader(),
                new P07T2Grader(),
                new P07T3Grader(),
                new P07T4Grader(),
                new P07T5Grader(),
                new P07T6Grader()
            };

            _project08Graders = new List<ITaskGrader>
            {
                new P08T1Grader(),
                new P08T2Grader(),
                new P08T3Grader(),
                new P08T4Grader(),
                new P08T5Grader(),
                new P08T6Grader()
            };

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

        public async Task<GradingResult> GradeProject01Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P01",
                ProjectName = "Morning Bean Coffee Sales"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project01Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject02Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P02",
                ProjectName = "Insurance Policy"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project02Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject03Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P03",
                ProjectName = "Munson Recipes"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project03Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject04Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P04",
                ProjectName = "Class schedule"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project04Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject05Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P05",
                ProjectName = "Book Purchases"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project05Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject06Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P06",
                ProjectName = "Sale Summary"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project06Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject07Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P07",
                ProjectName = "Tea Sales Report"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project07Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject08Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P08",
                ProjectName = "Book Sales"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project08Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }

        public async Task<GradingResult> GradeProject09Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P09",
                ProjectName = "Sales and Orders Report"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project09Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet, studentSheet));
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
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
            }

            return result;
        }
    }
}
