using OfficeOpenXml;
using MOS.ExcelGrading.Core.Graders.Project01;
using MOS.ExcelGrading.Core.Graders.Word.Project01;
using MOS.ExcelGrading.Core.Graders.Word.Project02;
using MOS.ExcelGrading.Core.Graders.Word;
using System.IO.Compression;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Graders.Project02;
using MOS.ExcelGrading.Core.Graders.Project03;
using MOS.ExcelGrading.Core.Graders.Project04;
using MOS.ExcelGrading.Core.Graders.Project05;
using MOS.ExcelGrading.Core.Graders.Project06;
using MOS.ExcelGrading.Core.Graders.Project07;
using MOS.ExcelGrading.Core.Graders.Project08;
using MOS.ExcelGrading.Core.Graders.Project09;
using MOS.ExcelGrading.Core.Graders.Project10;
using MOS.ExcelGrading.Core.Graders.Project11;
using MOS.ExcelGrading.Core.Graders.Project12;
using MOS.ExcelGrading.Core.Graders.Project13;
using MOS.ExcelGrading.Core.Graders.Project14;
using MOS.ExcelGrading.Core.Graders.Project15;
using MOS.ExcelGrading.Core.Graders.Project16;
using MOS.ExcelGrading.Core.Graders.Project18;
using MOS.ExcelGrading.Core.Graders.Project20;
using MOS.ExcelGrading.Core.Graders.Project22;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Graders.Word.Project03;
using MOS.ExcelGrading.Core.Graders.Word.Project05;
using MOS.ExcelGrading.Core.Graders.Word.Project07;
using MOS.ExcelGrading.Core.Graders.Word.Project09;
using MOS.ExcelGrading.Core.Graders.Word.Project11;
using MOS.ExcelGrading.Core.Graders.Word.Project13;
using System.Text;

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
        private readonly List<ITaskGrader> _project10Graders;
        private readonly List<ITaskGrader> _project11Graders;
        private readonly List<ITaskGrader> _project12Graders;
        private readonly List<ITaskGrader> _project13Graders;
        private readonly List<ITaskGrader> _project14Graders;
        private readonly List<ITaskGrader> _project15Graders;
        private readonly List<ITaskGrader> _project16Graders;
        private readonly List<ITaskGrader> _project18Graders;
        private readonly List<ITaskGrader> _project20Graders;
        private readonly List<ITaskGrader> _project22Graders;
        private readonly Dictionary<int, List<IWordTaskGrader>> _wordProjectGraders;
        private const decimal StandardProjectMaxScore = 125m;

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

            _project10Graders = new List<ITaskGrader>
            {
                new P10T1Grader(),
                new P10T2Grader(),
                new P10T3Grader(),
                new P10T4Grader(),
                new P10T5Grader(),
                new P10T6Grader()
            };

            _project11Graders = new List<ITaskGrader>
            {
                new P11T1Grader(),
                new P11T2Grader(),
                new P11T3Grader(),
                new P11T4Grader(),
                new P11T5Grader(),
                new P11T6Grader()
            };

            _project12Graders = new List<ITaskGrader>
            {
                new P12T1Grader(),
                new P12T2Grader(),
                new P12T3Grader(),
                new P12T4Grader(),
                new P12T5Grader(),
                new P12T6Grader()
            };

            _project13Graders = new List<ITaskGrader>
            {
                new P13T1Grader(),
                new P13T2Grader(),
                new P13T3Grader(),
                new P13T4Grader(),
                new P13T5Grader(),
                new P13T6Grader()
            };

            _project14Graders = new List<ITaskGrader>
            {
                new P14T1Grader(),
                new P14T2Grader(),
                new P14T3Grader(),
                new P14T4Grader(),
                new P14T5Grader(),
                new P14T6Grader()
            };

            _project15Graders = new List<ITaskGrader>
            {
                new P15T1Grader(),
                new P15T2Grader(),
                new P15T3Grader(),
                new P15T4Grader(),
                new P15T5Grader(),
                new P15T6Grader()
            };

            _project16Graders = new List<ITaskGrader>
            {
                new P16T1Grader(),
                new P16T2Grader(),
                new P16T3Grader(),
                new P16T4Grader(),
                new P16T5Grader(),
                new P16T6Grader()
            };

            _project18Graders = new List<ITaskGrader>
            {
                new P18T1Grader(),
                new P18T2Grader(),
                new P18T3Grader(),
                new P18T4Grader(),
                new P18T5Grader(),
                new P18T6Grader()
            };

            _project20Graders = new List<ITaskGrader>
            {
                new P20T1Grader(),
                new P20T2Grader(),
                new P20T3Grader(),
                new P20T4Grader(),
                new P20T5Grader(),
                new P20T6Grader()
            };

            _project22Graders = new List<ITaskGrader>
            {
                new P22T1Grader(),
                new P22T2Grader(),
                new P22T3Grader(),
                new P22T4Grader(),
                new P22T5Grader(),
                new P22T6Grader()
            };

            _wordProjectGraders = Enumerable.Range(1, 24)
                .ToDictionary(
                    projectNumber => projectNumber,
                    projectNumber => new List<IWordTaskGrader>
                    {
                        new WordProjectSkeletonTaskGrader(projectNumber)
                    });

            _wordProjectGraders[1] = new List<IWordTaskGrader>
            {
                new WP01T1Grader(),
                new WP01T2Grader(),
                new WP01T3Grader(),
                new WP01T4Grader(),
                new WP01T5Grader(),
                new WP01T6Grader()
            };

            _wordProjectGraders[2] = new List<IWordTaskGrader>
            {
                new WP02T1Grader()
            };

            _wordProjectGraders[3] = new List<IWordTaskGrader>
            {
                new WP03T1Grader(),
                new WP03T2Grader(),
                new WP03T3Grader(),
                new WP03T4Grader(),
                new WP03T5Grader(),
                new WP03T6Grader()
            };

            _wordProjectGraders[5] = new List<IWordTaskGrader>
            {
                new WP05T1Grader(),
                new WP05T2Grader(),
                new WP05T3Grader(),
                new WP05T4Grader(),
                new WP05T5Grader(),
                new WP05T6Grader(),
                new WP05T7Grader()
            };

            _wordProjectGraders[7] = new List<IWordTaskGrader>
            {
                new WP07T1Grader()
            };

            _wordProjectGraders[9] = new List<IWordTaskGrader>
            {
                new WP09T1Grader(),
                new WP09T2Grader(),
                new WP09T3Grader(),
                new WP09T4Grader(),
                new WP09T5Grader(),
                new WP09T6Grader()
            };

            _wordProjectGraders[11] = CreateWordProject11Graders();
            _wordProjectGraders[13] = CreateWordProject13LogicGraders(13);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private static List<IWordTaskGrader> CreateWordProject11Graders()
        {
            return new List<IWordTaskGrader>
            {
                new WP11T1Grader(),
                new WP11T2Grader(),
                new WP11T3Grader(),
                new WP11T4Grader(),
                new WP11T5Grader(),
                new WP11T6Grader()
            };
        }

        private static List<IWordTaskGrader> CreateWordProject13LogicGraders(int projectNumber)
        {
            return new List<IWordTaskGrader>
            {
                new WP13T1Grader($"W{projectNumber:00}-T01"),
                new WP13T2Grader($"W{projectNumber:00}-T02"),
                new WP13T3Grader($"W{projectNumber:00}-T03"),
                new WP13T4Grader($"W{projectNumber:00}-T04"),
                new WP13T5Grader($"W{projectNumber:00}-T05"),
                new WP13T6Grader($"W{projectNumber:00}-T06")
            };
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject07Async(Stream studentFile, string? sourceFileName = null)
        {
            if (IsPlainTextWordInput(7, sourceFileName))
            {
                return await GradeWordProjectAsync(7, studentFile, sourceFileName);
            }

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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
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
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject10Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P10",
                ProjectName = "Bellows Institute"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project10Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject11Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P11",
                ProjectName = "Toy Store Report"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project11Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject12Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P12",
                ProjectName = "Clothing Orders"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project12Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject13Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P13",
                ProjectName = "Retreat Plans"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project13Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject14Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P14",
                ProjectName = "Policy Renewals"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project14Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject15Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P15",
                ProjectName = "Tailspins Data Report"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project15Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject16Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P16",
                ProjectName = "List of Product"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project16Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject18Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P18",
                ProjectName = "Bank Accounts"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project18Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject20Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P20",
                ProjectName = "Air Miles"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project20Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeProject22Async(Stream studentFile)
        {
            var result = new GradingResult
            {
                ProjectId = "P22",
                ProjectName = "Student Scores"
            };

            try
            {
                using var studentPackage = new ExcelPackage(studentFile);
                var studentSheet = studentPackage.Workbook.Worksheets[0];

                foreach (var grader in _project22Graders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(studentSheet));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        public async Task<GradingResult> GradeWordProjectAsync(int projectNumber, Stream studentFile, string? sourceFileName = null)
        {
            var result = new GradingResult
            {
                ProjectId = $"W{projectNumber:00}",
                ProjectName = projectNumber == 11 ? "River Cruises" : $"Word Project {projectNumber:00}"
            };

            if (projectNumber < 1 || projectNumber > 24)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Project Word {projectNumber:00} khong duoc ho tro." }
                });
                ApplyProjectScoringModel(result);
                return result;
            }

            if (!_wordProjectGraders.TryGetValue(projectNumber, out var projectGraders) || projectGraders.Count == 0)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string>
                    {
                        $"Chua cau hinh grader Word cho project {projectNumber:00}."
                    }
                });
                ApplyProjectScoringModel(result);
                return result;
            }

            try
            {
                var context = await LoadWordGradingContextAsync(projectNumber, studentFile, sourceFileName);
                foreach (var grader in projectGraders)
                {
                    var taskResult = await Task.Run(() => grader.Grade(context));
                    result.TaskResults.Add(taskResult);
                }

                ApplyProjectScoringModel(result);
            }
            catch (Exception ex)
            {
                result.TaskResults.Add(new TaskResult
                {
                    TaskId = "ERROR",
                    TaskName = "Loi he thong",
                    Errors = new List<string> { $"Loi: {ex.Message}" }
                });
                ApplyProjectScoringModel(result);
            }

            return result;
        }

        private static async Task<WordGradingContext> LoadWordGradingContextAsync(
            int projectNumber,
            Stream studentFile,
            string? sourceFileName)
        {
            if (studentFile == null)
            {
                throw new InvalidOperationException("Khong co file dau vao de cham Word.");
            }

            await using var memoryStream = new MemoryStream();
            await studentFile.CopyToAsync(memoryStream);
            memoryStream.Position = 0;

            if (IsPlainTextWordInput(projectNumber, sourceFileName))
            {
                return await LoadWordPlainTextContextAsync(projectNumber, memoryStream, sourceFileName);
            }

            try
            {
                using var zipArchive = new ZipArchive(memoryStream, ZipArchiveMode.Read, leaveOpen: true);
                var entryNames = zipArchive.Entries
                    .Select(entry => entry.FullName.Replace("\\", "/", StringComparison.Ordinal))
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                var xmlParts = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
                foreach (var entry in zipArchive.Entries)
                {
                    var normalizedName = entry.FullName.Replace("\\", "/", StringComparison.Ordinal);
                    if (!normalizedName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    try
                    {
                        using var entryStream = entry.Open();
                        xmlParts[normalizedName] = XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
                    }
                    catch
                    {
                        // Ignore malformed or unsupported XML parts to keep grading resilient.
                    }
                }

                var documentRelationshipsXml = TryLoadXmlPart(zipArchive, "word/_rels/document.xml.rels");
                var documentRelationships = BuildWordRelationships(documentRelationshipsXml);

                xmlParts.TryGetValue("word/document.xml", out var mainDocumentXml);
                xmlParts.TryGetValue("docProps/core.xml", out var corePropertiesXml);
                xmlParts.TryGetValue("word/numbering.xml", out var numberingXml);
                xmlParts.TryGetValue("word/_rels/document.xml.rels", out var relationshipsXmlFromParts);
                documentRelationshipsXml ??= relationshipsXmlFromParts;

                return new WordGradingContext
                {
                    ProjectNumber = projectNumber,
                    SourceFileName = sourceFileName ?? $"project{projectNumber:00}.docx",
                    PackageBytes = memoryStream.ToArray(),
                    HasMainDocumentPart = entryNames.Contains("word/document.xml"),
                    PartCount = entryNames.Count,
                    Entries = entryNames,
                    MainDocumentXml = mainDocumentXml,
                    CorePropertiesXml = corePropertiesXml,
                    NumberingXml = numberingXml,
                    DocumentRelationshipsXml = documentRelationshipsXml,
                    DocumentRelationships = documentRelationships,
                    XmlParts = xmlParts
                };
            }
            catch (InvalidDataException ex)
            {
                throw new InvalidOperationException("File Word khong dung dinh dang OpenXML (.docx).", ex);
            }
        }

        private static bool IsPlainTextWordInput(int projectNumber, string? sourceFileName)
        {
            if (projectNumber != 7)
            {
                return false;
            }

            var extension = Path.GetExtension(sourceFileName ?? string.Empty);
            return string.Equals(extension, ".txt", StringComparison.OrdinalIgnoreCase);
        }

        private static async Task<WordGradingContext> LoadWordPlainTextContextAsync(
            int projectNumber,
            MemoryStream memoryStream,
            string? sourceFileName)
        {
            memoryStream.Position = 0;
            using var reader = new StreamReader(memoryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
            var text = await reader.ReadToEndAsync();

            var plainTextXml = BuildPlainTextWordDocumentXml(text);
            var entries = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "plain-text-input"
            };

            return new WordGradingContext
            {
                ProjectNumber = projectNumber,
                SourceFileName = sourceFileName ?? $"project{projectNumber:00}.txt",
                PackageBytes = memoryStream.ToArray(),
                HasMainDocumentPart = true,
                PartCount = 1,
                Entries = entries,
                MainDocumentXml = plainTextXml,
                CorePropertiesXml = null,
                NumberingXml = null,
                DocumentRelationshipsXml = null,
                DocumentRelationships = new Dictionary<string, WordDocumentRelationship>(StringComparer.OrdinalIgnoreCase),
                XmlParts = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase)
            };
        }

        private static XDocument BuildPlainTextWordDocumentXml(string text)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var lines = (text ?? string.Empty)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Split('\n');

            var paragraphs = lines.Select(line =>
                new XElement(w + "p",
                    new XElement(w + "r",
                        new XElement(w + "t", line))));

            return new XDocument(
                new XElement(w + "document",
                    new XElement(w + "body", paragraphs)));
        }

        private static XDocument? TryLoadXmlPart(ZipArchive archive, string entryName)
        {
            var entry = archive.GetEntry(entryName);
            if (entry == null)
            {
                return null;
            }

            using var entryStream = entry.Open();
            return XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
        }

        private static Dictionary<string, WordDocumentRelationship> BuildWordRelationships(XDocument? documentRelationshipsXml)
        {
            var relationships = new Dictionary<string, WordDocumentRelationship>(StringComparer.OrdinalIgnoreCase);
            if (documentRelationshipsXml?.Root == null)
            {
                return relationships;
            }

            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            foreach (var relationNode in documentRelationshipsXml.Root.Elements(relNs + "Relationship"))
            {
                var relationshipId = relationNode.Attribute("Id")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(relationshipId))
                {
                    continue;
                }

                relationships[relationshipId] = new WordDocumentRelationship
                {
                    Id = relationshipId,
                    Type = relationNode.Attribute("Type")?.Value ?? string.Empty,
                    Target = relationNode.Attribute("Target")?.Value ?? string.Empty
                };
            }

            return relationships;
        }

        private static void ApplyProjectScoringModel(GradingResult result)
        {
            var gradableTasks = result.TaskResults
                .Where(task => !string.Equals(task.TaskId, "ERROR", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (gradableTasks.Count == 0)
            {
                result.TotalScore = 0m;
                result.MaxScore = StandardProjectMaxScore;
                EnsureCorrectionGuidance(result);
                return;
            }

            var totalSourceMax = gradableTasks.Sum(task => task.MaxScore > 0m ? task.MaxScore : 1m);
            if (totalSourceMax <= 0m)
            {
                totalSourceMax = gradableTasks.Count;
            }

            decimal allocatedMax = 0m;
            for (var i = 0; i < gradableTasks.Count; i++)
            {
                var task = gradableTasks[i];
                var sourceMax = task.MaxScore > 0m ? task.MaxScore : 1m;
                var isLastTask = i == gradableTasks.Count - 1;

                var scaledMax = isLastTask
                    ? StandardProjectMaxScore - allocatedMax
                    : Math.Round(StandardProjectMaxScore * sourceMax / totalSourceMax, 2, MidpointRounding.AwayFromZero);

                if (scaledMax < 0m)
                {
                    scaledMax = 0m;
                }

                allocatedMax += scaledMax;

                var completionRatio = task.MaxScore > 0m
                    ? Math.Clamp(task.Score / task.MaxScore, 0m, 1m)
                    : 0m;

                var scaledScore = Math.Round(scaledMax * completionRatio, 2, MidpointRounding.AwayFromZero);
                if (scaledScore > scaledMax)
                {
                    scaledScore = scaledMax;
                }

                task.MaxScore = scaledMax;
                task.Score = scaledScore;
            }

            result.MaxScore = StandardProjectMaxScore;
            result.TotalScore = Math.Round(gradableTasks.Sum(task => task.Score), 2, MidpointRounding.AwayFromZero);
            if (result.TotalScore > result.MaxScore)
            {
                result.TotalScore = result.MaxScore;
            }

            EnsureCorrectionGuidance(result);
        }

        private static void EnsureCorrectionGuidance(GradingResult result)
        {
            foreach (var task in result.TaskResults)
            {
                if (task.Errors.Count > 0 && task.FixActions.Count == 0)
                {
                    task.FixActions.Add("Xem lại các lỗi được liệt kê trong Errors và chỉnh lại tài liệu theo đúng yêu cầu của task.");
                }
            }
        }
    }
}
