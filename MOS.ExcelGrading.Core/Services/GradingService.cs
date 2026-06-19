using OfficeOpenXml;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project01;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project01;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project02;
using MOS.ExcelGrading.Core.Graders.OTTH.Word;
using System.IO.Compression;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project02;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project03;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project04;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project05;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project06;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project07;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project08;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project09;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project10;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project11;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project12;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project13;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project14;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project15;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project16;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project18;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project20;
using MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project22;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project03;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project05;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project07;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project09;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project11;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project13;
using MOS.ExcelGrading.Core.Graders.OTTH.Word.Project15;
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
            _wordProjectGraders[15] = new List<IWordTaskGrader>
            {
                new WP15T1Grader()
            };

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

        public List<ExamPublicationTaskSnapshotItemDto> GetTaskSnapshotForEndpoint(string gradingApiEndpoint)
        {
            if (string.IsNullOrWhiteSpace(gradingApiEndpoint))
            {
                return new List<ExamPublicationTaskSnapshotItemDto>();
            }

            var normalizedEndpoint = GradingApiEndpoints.NormalizeEndpoint(gradingApiEndpoint);
            if (!GradingApiEndpoints.TryExtractSubject(normalizedEndpoint, out var subject) ||
                !GradingApiEndpoints.TryExtractProjectNumber(normalizedEndpoint, out var projectNumber))
            {
                return new List<ExamPublicationTaskSnapshotItemDto>();
            }

            if (string.Equals(subject, AssignmentFileSubjects.Word, StringComparison.OrdinalIgnoreCase))
            {
                return _wordProjectGraders.TryGetValue(projectNumber, out var wordGraders)
                    ? MapTaskSnapshot(wordGraders)
                    : new List<ExamPublicationTaskSnapshotItemDto>();
            }

            if (!string.Equals(subject, AssignmentFileSubjects.Excel, StringComparison.OrdinalIgnoreCase))
            {
                return new List<ExamPublicationTaskSnapshotItemDto>();
            }

            return GetExcelTaskSnapshot(projectNumber);
        }

        private List<ExamPublicationTaskSnapshotItemDto> GetExcelTaskSnapshot(int projectNumber)
        {
            return projectNumber switch
            {
                1 => MapTaskSnapshot(_project01Graders),
                2 => MapTaskSnapshot(_project02Graders),
                3 => MapTaskSnapshot(_project03Graders),
                4 => MapTaskSnapshot(_project04Graders),
                5 => MapTaskSnapshot(_project05Graders),
                6 => MapTaskSnapshot(_project06Graders),
                7 => MapTaskSnapshot(_project07Graders),
                8 => MapTaskSnapshot(_project08Graders),
                9 => MapTaskSnapshot(_project09Graders),
                10 => MapTaskSnapshot(_project10Graders),
                11 => MapTaskSnapshot(_project11Graders),
                12 => MapTaskSnapshot(_project12Graders),
                13 => MapTaskSnapshot(_project13Graders),
                14 => MapTaskSnapshot(_project14Graders),
                15 => MapTaskSnapshot(_project15Graders),
                16 => MapTaskSnapshot(_project16Graders),
                18 => MapTaskSnapshot(_project18Graders),
                20 => MapTaskSnapshot(_project20Graders),
                22 => MapTaskSnapshot(_project22Graders),
                _ => new List<ExamPublicationTaskSnapshotItemDto>()
            };
        }

        private static List<ExamPublicationTaskSnapshotItemDto> MapTaskSnapshot(IEnumerable<ITaskGrader> graders)
        {
            return graders
                .Select(grader => new ExamPublicationTaskSnapshotItemDto
                {
                    TaskId = NormalizeTaskValue(grader.TaskId),
                    TaskName = NormalizeTaskValue(TaskNameFormatter.RemovePrefix(grader.TaskName)),
                    MaxScore = (double)grader.MaxScore
                })
                .ToList();
        }

        private static List<ExamPublicationTaskSnapshotItemDto> MapTaskSnapshot(IEnumerable<IWordTaskGrader> graders)
        {
            return graders
                .Select(grader => new ExamPublicationTaskSnapshotItemDto
                {
                    TaskId = NormalizeTaskValue(grader.TaskId),
                    TaskName = NormalizeTaskValue(TaskNameFormatter.RemovePrefix(grader.TaskName)),
                    MaxScore = (double)grader.MaxScore
                })
                .ToList();
        }

        private static string? NormalizeTaskValue(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
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
                    TaskName = "Lỗi hệ thống",
                    Errors = new List<string> { $"Lỗi: {ex.Message}" }
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

                NormalizeTaskFeedbackLines(result);
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

        private static void NormalizeTaskFeedbackLines(GradingResult result)
        {
            foreach (var task in result.TaskResults)
            {
                task.Errors = NormalizeFeedbackLines(task.Errors);
                task.FixActions = new SingleFixActionList(NormalizeFeedbackLines(task.FixActions));
            }
        }

        private static List<string> NormalizeFeedbackLines(List<string>? lines)
        {
            if (lines == null || lines.Count == 0)
            {
                return new List<string>();
            }

            var normalized = lines
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            return normalized;
        }

        private static void EnsureTaskErrorContract(GradingResult result)
        {
            foreach (var task in result.TaskResults)
            {
                if (task.IsPassed || task.Errors.Count > 0)
                {
                    continue;
                }

                task.Errors.Add("Bài làm chưa đạt yêu cầu của task này.");
            }
        }

        private static void BuildTaskDisplayIssues(GradingResult result)
        {
            foreach (var task in result.TaskResults)
            {
                task.DisplayIssues = new List<TaskDisplayIssue>();

                if (task.Errors.Count == 0)
                {
                    continue;
                }

                var heading = string.IsNullOrWhiteSpace(task.TaskName)
                    ? task.TaskId?.Trim() ?? string.Empty
                    : task.TaskName.Trim();

                for (var index = 0; index < task.Errors.Count; index++)
                {
                    var message = task.Errors[index]?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(message))
                    {
                        continue;
                    }

                    var fixAction = task.FixActions.Count > index
                        ? task.FixActions[index]
                        : task.FixActions.FirstOrDefault() ?? string.Empty;

                    task.DisplayIssues.Add(new TaskDisplayIssue
                    {
                        Heading = heading,
                        Message = message,
                        FixAction = fixAction?.Trim() ?? string.Empty
                    });
                }
            }
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
            NormalizeTaskFeedbackLines(result);

            var gradableTasks = result.TaskResults
                .Where(task => !string.Equals(task.TaskId, "ERROR", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (gradableTasks.Count == 0)
            {
                result.TotalScore = 0m;
                result.MaxScore = StandardProjectMaxScore;
                EnsureCorrectionGuidance(result);
                EnsureTaskErrorContract(result);
                NormalizeTaskFeedbackLines(result);
                BuildTaskDisplayIssues(result);
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
            EnsureTaskErrorContract(result);
            NormalizeTaskFeedbackLines(result);
            BuildTaskDisplayIssues(result);
        }

        private static void EnsureCorrectionGuidance(GradingResult result)
        {
            foreach (var task in result.TaskResults)
            {
                if (task.Errors.Count > 0 && task.FixActions.Count == 0)
                {
                    var fixAction = IsExcelTask(task)
                        ? BuildExcelFixAction(task)
                        : "Xem lại các lỗi được liệt kê trong Errors và chỉnh lại tài liệu theo đúng yêu cầu của task.";

                    task.FixActions.Add(fixAction);
                }
            }
        }

        private static bool IsExcelTask(TaskResult task)
        {
            return task.TaskId.StartsWith("P", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildExcelFixAction(TaskResult task)
        {
            var taskId = task.TaskId?.Trim().ToUpperInvariant() ?? string.Empty;
            var taskName = TaskNameFormatter.RemovePrefix(task.TaskName);

            if (ExcelFixActions.TryGetValue(taskId, out var fixAction))
            {
                return fixAction;
            }

            if (string.IsNullOrWhiteSpace(taskName))
            {
                return "Mở workbook Excel, tìm đúng worksheet hoặc vùng dữ liệu đang bị báo lỗi, thực hiện lại đúng thao tác theo đề bài, rồi lưu file và chấm lại.";
            }

            return $"Mở worksheet hoặc vùng dữ liệu của \"{taskName}\", thực hiện lại đúng thao tác Excel theo đề bài, rồi lưu file và chấm lại.";
        }

        private static readonly Dictionary<string, string> ExcelFixActions = new(StringComparer.OrdinalIgnoreCase)
        {
            ["P01-T01"] = "Mở sheet Documentation, dùng Format Painter/copy định dạng vùng A1:A2, chuyển sang sheet Menu Items, áp dụng đúng định dạng đó cho tiêu đề/phụ đề tương ứng, lưu file rồi chấm lại.",
            ["P01-T02"] = "Mở sheet Menu Items, chọn bảng Table2, vào Table Design > Table Name, đổi tên bảng thành Units_Sold, nhấn Enter, lưu file rồi chấm lại.",
            ["P01-T03"] = "Mở ô C48, nhập công thức SUM cộng đúng 4 named range được yêu cầu, kiểm tra kết quả tổng, lưu file rồi chấm lại.",
            ["P01-T04"] = "Mở ô K48, nhập công thức đếm số mục thiếu của tháng 9 theo đúng vùng dữ liệu yêu cầu, kiểm tra kết quả đếm, lưu file rồi chấm lại.",
            ["P01-T05"] = "Mở các cột % Change của Units Sold và Gross Sales, nhập công thức phần trăm thay đổi đúng cho dòng đầu tiên, fill xuống toàn bộ vùng yêu cầu, giữ định dạng phần trăm, lưu file rồi chấm lại.",

            ["P02-T01"] = "Mở sheet New Policy, chọn cột Agent trong bảng, căn trái nội dung và đặt thụt lề đúng yêu cầu, lưu file rồi chấm lại.",
            ["P02-T02"] = "Mở sheet New Policy, chọn vùng dữ liệu nguồn theo tháng, vào Insert > Sparklines > Win/Loss, đặt Location Range là J5:J13, lưu file rồi chấm lại.",
            ["P02-T03"] = "Mở bảng New Policy, bật Total Row trong Table Design, chọn hàm tổng phù hợp cho các cột theo tháng, lưu file rồi chấm lại.",
            ["P02-T04"] = "Mở cột Inactive months, nhập công thức đếm số tháng không có policy mới cho từng dòng, fill xuống hết cột, lưu file rồi chấm lại.",
            ["P02-T05"] = "Mở cột Email, tạo email bằng cách ghép First name với @humongousinsurance.com, fill công thức xuống toàn bộ cột, lưu file rồi chấm lại.",
            ["P02-T06"] = "Mở sheet New Policy,Chọn biểu đồ, vào Chart Design > Quick Layout, chọn Layout 3, lưu file rồi chấm lại.",
            ["P02-T07"] = "Chọn biểu đồ, vào Chart Design > Change Colors, chọn Colorful Palette 2, kiểm tra màu biểu đồ đã đổi, lưu file rồi chấm lại.",

            ["P03-T01"] = "Mở sheet Ingredients, chọn vùng A1:N1, dùng Merge & Center hoặc Merge Cells để gộp thành một ô, lưu file rồi chấm lại.",
            ["P03-T02"] = "Mở sheet Ingredients, chọn các cột A:N, dùng AutoFit Column Width để tự động vừa nội dung, lưu file rồi chấm lại.",
            ["P03-T03"] = "Mở Page Setup/Header Footer, đặt phần Header bên phải là Sequential, sau đó chuyển workbook về chế độ xem Normal, lưu file rồi chấm lại.",
            ["P03-T04"] = "Mở ô A6, tạo hyperlink trỏ đến sheet Description ô A18, kiểm tra liên kết nội bộ đúng đích, lưu file rồi chấm lại.",
            ["P03-T05"] = "Mở Page Setup, đặt Orientation là Landscape và Scaling là Fit all columns on one page, lưu file rồi chấm lại.",
            ["P03-T06"] = "Mở tùy chọn Quick Access Toolbar, thêm lệnh Quick Print vào thanh công cụ, lưu workbook rồi chấm lại.",

            ["P04-T01"] = " import dữ liệu Substitutes vào đúng sheet/vị trí yêu cầu, chuyển vùng dữ liệu thành bảng và áp dụng Table Style Medium 1, lưu file rồi chấm lại.",
            ["P04-T02"] = "Mở sheet Number of course hours, chọn cột B:G, đặt Column Width bằng 12, lưu file rồi chấm lại.",
            ["P04-T03"] = "Mở sheet Enrollment, chọn dữ liệu nguồn, vào Insert > Sparklines > Line, đặt Location Range là G5:G25, lưu file rồi chấm lại.",
            ["P04-T04"] = "Mở sheet Classes, chọn bảng, vào Table Design > Convert to Range, xác nhận chuyển thành range và giữ định dạng, lưu file rồi chấm lại.",
            ["P04-T05"] = "Chọn biểu đồ, vào Chart Design > Move Chart, chọn New sheet và đặt tên Graduation Chart, lưu file rồi chấm lại.",
            ["P04-T06"] = "Chọn biểu đồ, thêm hoặc sửa tiêu đề trục dọc chính thành Hours, lưu file rồi chấm lại.",
            ["P04-T07"] = "Mở sheet Classes, dùng Sort nhiều cấp: Instructor A-Z trước, sau đó Section giảm dần, lưu file rồi chấm lại.",

            ["P05-T01"] = "Mở File > Info > Properties, đặt Company thành Salon International, lưu file rồi chấm lại.",
            ["P05-T02"] = "Mở cột Difference, nhập công thức Selling Price trừ Cost cho dòng đầu tiên, fill xuống toàn bộ cột và giữ nguyên định dạng, lưu file rồi chấm lại.",
            ["P05-T03"] = "Mở ô F37, nhập công thức tính trung bình Selling Price cho các dòng có nhà cung cấp \"Fabrikam\", Inc., lưu file rồi chấm lại.",
            ["P05-T04"] = " kéo hoặc Move sheet Annual Purchases đến vị trí giữa sheet Works và sheet Titles, lưu file rồi chấm lại.",
            ["P05-T05"] = "Mở sheet Works, chọn hình ảnh, đặt Rotation về 0 độ trong Format Picture/Size Properties, lưu file rồi chấm lại.",
            ["P05-T06"] = "Mở cửa sổ workbook thứ hai, dùng View Side by Side và Arrange All theo kiểu trên-dưới, lưu file rồi chấm lại.",

            ["P06-T01"] = "Mở vùng F4:F11, tạo Conditional Formatting Greater Than 5000000 với Yellow Fill và Dark Yellow Text, lưu file rồi chấm lại.",
            ["P06-T02"] = "Mở Region 1, dùng Sort nhiều cấp: Product A-Z trước, sau đó Total sales giảm dần, lưu file rồi chấm lại.",
            ["P06-T03"] = "Mở sheet Forecasts, tại cột Quarter 2 nhập công thức Quarter 1 nhân named range Q2_Increase, fill xuống các dòng yêu cầu, lưu file rồi chấm lại.",
            ["P06-T04"] = "Mở sheet Summary ô B15, nhập hàm MAX cho cột Total sales, kiểm tra kết quả lớn nhất, lưu file rồi chấm lại.",
            ["P06-T05"] = "Chọn biểu đồ Comparison, vào Chart Design > Switch Row/Column, kiểm tra series/trục đã hoán đổi, lưu file rồi chấm lại.",
            ["P06-T06"] = " chọn pie chart trên sheet Qtr 2, vào Move Chart, chọn New sheet và đặt tên Qtr 2 Chart, lưu file rồi chấm lại.",

            ["P07-T01"] = "Mở sheet Drinks, import Drinks.txt bắt đầu tại ô A7, bật Use first row as headers khi tạo bảng, lưu file rồi chấm lại.",
            ["P07-T02"] = " chọn bảng Tea, vào Table Design > Table Styles, chọn Blue, Table Style Medium 9, lưu file rồi chấm lại.",
            ["P07-T03"] = "Chọn biểu đồ Tea, áp dụng Quick Layout 9, đặt Vertical Axis Title là Price và xóa Horizontal Axis Title, lưu file rồi chấm lại.",
            ["P07-T04"] = "Mở sheet Total Cookie Sales, chọn A3:A8, áp dụng Pattern Fill đúng mẫu/màu theo yêu cầu, lưu file rồi chấm lại.",
            ["P07-T05"] = "Mở ô B3 trên Total Cookie Sales, nhập công thức SUM(Table2[Chocolate Mint Chip]), lưu file rồi chấm lại.",
            ["P07-T06"] = " copy vùng Q1 Sales A4:E9, sang sheet Seedling Sales ô A4, Paste Special > Transpose, lưu file rồi chấm lại.",

            ["P08-T01"] = "Mở sheet Summary ô A2, tạo hyperlink đến www.nodpublishers.com và nhập đúng ScreenTip theo yêu cầu, lưu file rồi chấm lại.",
            ["P08-T02"] = "Mở sheet Sale History, bật Show Formulas trong tab Formulas hoặc View, lưu file rồi chấm lại.",
            ["P08-T03"] = "Mở File > Info > Check for Issues > Inspect Document, xóa thông tin cá nhân trong Document Properties, lưu file rồi chấm lại.",
            ["P08-T04"] = "Mở cột Authors Premium, nhập công thức IF(Books sold > 10000, 500, 100), fill xuống toàn bộ vùng yêu cầu, lưu file rồi chấm lại.",
            ["P08-T05"] = "Mở cột Sales Postal Code, dùng UPPER kết hợp lấy 3 ký tự đầu của giá trị nguồn, fill xuống toàn bộ cột, lưu file rồi chấm lại.",
            ["P08-T06"] = "Chọn biểu đồ Summary,Mở rộng vùng dữ liệu chart để bao gồm Current Year, kiểm tra series mới đã xuất hiện, lưu file rồi chấm lại.",

            ["P09-T01"] = "Chọn biểu đồ, đặt Pattern Fill 10% cho Plot Area và Pattern Fill 50% cho Chart Area theo màu yêu cầu, lưu file rồi chấm lại.",
            ["P09-T02"] = "Mở ô A1, bỏ merge, áp dụng Cell Style Title, đặt font size 24 và Bold, lưu file rồi chấm lại.",
            ["P09-T03"] = "Chọn biểu đồ, hiển thị Legend ở bên phải và bật tùy chọn cho phép legend overlay/tràn lên chart nếu được yêu cầu, lưu file rồi chấm lại.",
            ["P09-T04"] = "Mở bảng dữ liệu, lọc cột Total với điều kiện từ 34000 đến 45000, lưu file rồi chấm lại.",
            ["P09-T05"] = "Mở dữ liệu, dùng Subtotal theo Shirt Color, bật Page break between groups và Summary below data/Grand Total theo yêu cầu, lưu file rồi chấm lại.",
            ["P09-T06"] = "Mở sheet Farmers & Market, chọn vùng dữ liệu yêu cầu, vào Insert > Pie Chart > 3-D Pie, đặt biểu đồ trên cùng sheet Farmers & Market đúng vị trí/kích thước yêu cầu, lưu file rồi chấm lại.",

            ["P10-T01"] = "Mở sheet Last semester, chọn A3:F3, bật Wrap Text, lưu file rồi chấm lại.",
            ["P10-T02"] = "Mở sheet Enrollment summary, chọn đúng vùng dữ liệu Enrollment, vào Formulas > Define Name và đặt tên Enrollment, lưu file rồi chấm lại.",
            ["P10-T03"] = "Mở sheet Income, chọn A3:B7, tạo Table có header và áp dụng style Light 14, lưu file rồi chấm lại.",
            ["P10-T04"] = "Mở sheet Last semester, trong bảng xóa dòng có Agriculture bằng thao tác Delete Table Rows, lưu file rồi chấm lại.",
            ["P10-T05"] = "Mở sheet Next semester, chọn dữ liệu Program và Average cost, chèn biểu đồ Clustered Column, đặt vị trí/kích thước đúng yêu cầu, lưu file rồi chấm lại.",
            ["P10-T06"] = "Chọn biểu đồ trên Enrollment summary, áp dụng Chart Style 7 và Change Colors > Monochromatic Palette 6, lưu file rồi chấm lại.",

            ["P11-T01"] = "Mở sheet Games, lần lượt chọn các vùng A12:B12 đến A18:B18 và merge từng dòng đúng yêu cầu, lưu file rồi chấm lại.",
            ["P11-T02"] = "Mở sheet Shareholders Info, tìm dòng Annual Report và đặt Row Height bằng 30, lưu file rồi chấm lại.",
            ["P11-T03"] = "Đổi tên worksheet Outdoor Toys thành Outdoor Sports, lưu file rồi chấm lại.",
            ["P11-T04"] = "Mở sheet Shareholders Info ô C5, tạo hyperlink đúng địa chỉ và đặt Display Text đúng yêu cầu, lưu file rồi chấm lại.",
            ["P11-T05"] = "Trên từng worksheet yêu cầu,Mở Page Setup và đặt Scaling/Fit Sheet on One Page, lưu file rồi chấm lại.",
            ["P11-T06"] = "Mở sheet Costs, vào Page Setup > Print Titles, đặt Rows to repeat at top là $1:$3, lưu file rồi chấm lại.",

            ["P12-T01"] = "Mở sheet Range, chọn E7:F7, merge thành một ô, lưu file rồi chấm lại.",
            ["P12-T02"] = "Mở sheet Prices ô A1, áp dụng Cell Style Title, lưu file rồi chấm lại.",
            ["P12-T03"] = "Mở sheet Orders, bật filter và lọc đúng khách hàng The House of Alpine Skiing, lưu file rồi chấm lại.",
            ["P12-T04"] = "Mở sheet Prices cột Tax, nhập công thức Unit price nhân L$2 với tham chiếu tuyệt đối đúng, fill xuống toàn bộ cột, lưu file rồi chấm lại.",
            ["P12-T05"] = "Mở sheet Prices cột Inventory Notice, nhập công thức IF kiểm tra tỷ lệ dưới 15% thì trả Low, ngược lại để trống, fill xuống, lưu file rồi chấm lại.",
            ["P12-T06"] = "Chọn Inventory chart, bật Chart Title và Data Labels, đặt Data Labels ở vị trí Outside End, lưu file rồi chấm lại.",

            ["P13-T01"] = "Mở sheet Shirt Orders, dùng Find/Replace thay Amber bằng Gold trong vùng yêu cầu, lưu file rồi chấm lại.",
            ["P13-T02"] = "Mở sheet Shirt Orders ô C2, nhập công thức SUMIF tính chi phí cho màu Blue theo đúng vùng dữ liệu, lưu file rồi chấm lại.",
            ["P13-T03"] = "Mở sheet Shirt Orders ô C3, nhập công thức COUNTIF đếm size Large theo đúng vùng dữ liệu, lưu file rồi chấm lại.",
            ["P13-T04"] = "Mở sheet Shirt Orders, thêm dòng Subtotal tại dòng 201 và nhập tổng phụ đúng tại D201/F201, lưu file rồi chấm lại.",
            ["P13-T05"] = "Mở sheet Attendees, chuyển sang Page Layout View và đặt page break/ngắt trang đúng vị trí yêu cầu, lưu file rồi chấm lại.",
            ["P13-T06"] = "Mở sheet Price List ô H5, nhập công thức dùng Structured Reference đúng theo bảng, fill nếu cần, lưu file rồi chấm lại.",

            ["P14-T01"] = "Mở sheet January, chọn vùng A4:F20, vào Page Layout > Print Area > Set Print Area, lưu file rồi chấm lại.",
            ["P14-T02"] = "Mở sheet March, bật filter và lọc cột Policy Type chỉ còn PM, lưu file rồi chấm lại.",
            ["P14-T03"] = "Mở sheet February cột Discount, nhập công thức Discount đúng theo yêu cầu, fill xuống toàn bộ cột, lưu file rồi chấm lại.",
            ["P14-T04"] = "Mở sheet February cột Policy Type, nhập công thức LEFT lấy 2 ký tự đầu theo yêu cầu, fill xuống toàn bộ cột, lưu file rồi chấm lại.",
            ["P14-T05"] = "Mở sheet Summary,Chọn biểu đồ, vào Format Chart Area > Alt Text và đặt mô tả/title là Renewal Data, lưu file rồi chấm lại.",
            ["P14-T06"] = "Mở sheet Sales cột Auction ID, nhập công thức RANDBETWEEN(1000,2000), fill xuống vùng yêu cầu, lưu file rồi chấm lại.",

            ["P15-T01"] = "Mở sheet Products cột Weight, đặt Number Format hiển thị 3 chữ số thập phân, lưu file rồi chấm lại.",
            ["P15-T02"] = "Mở sheet Products ô G3, nhập công thức SUMIF tính tổng cho Magic Supplies theo đúng vùng dữ liệu, lưu file rồi chấm lại.",
            ["P15-T03"] = "Mở sheet Orders, chọn vùng dữ liệu yêu cầu, tạo Conditional Formatting Above Average với green style, lưu file rồi chấm lại.",
            ["P15-T04"] = "Mở sheet Customers ô N5, nhập công thức COUNTIF đếm United States theo đúng vùng dữ liệu, lưu file rồi chấm lại.",
            ["P15-T05"] = "Mở sheet Customers cột CurrentAge, nhập/fill công thức tuổi hiện tại đúng và không thay đổi định dạng sẵn có, lưu file rồi chấm lại.",
            ["P15-T06"] = "Chọn các worksheet yêu cầu, đặt màu tab là Pink Accent 1, lưu file rồi chấm lại.",

            ["P16-T01"] = "Mở sheet Products, chọn ô dưới 2 hàng đầu rồi vào View > Freeze Panes để cố định 2 hàng đầu, lưu file rồi chấm lại.",
            ["P16-T02"] = "Mở sheet Products ô A1, đặt Horizontal Alignment là Left, lưu file rồi chấm lại.",
            ["P16-T03"] = "Mở sheet Products cột Quantity, tạo Conditional Formatting Icon Set dạng 3 traffic lights theo đúng ngưỡng yêu cầu, lưu file rồi chấm lại.",
            ["P16-T04"] = "Mở sheet Products, chọn bảng và áp dụng Table Style Medium 1, lưu file rồi chấm lại.",
            ["P16-T05"] = "Mở sheet Products ô F3, nhập công thức Estimated Value đúng, fill down đến hết dữ liệu, lưu file rồi chấm lại.",
            ["P16-T06"] = "Mở sheet Summary,Chọn biểu đồ và áp dụng Change Colors > Colorful Palette 2, lưu file rồi chấm lại.",

            ["P18-T01"] = "Chọn phạm vi cần đặt tên, vào Formulas > Define Name đặt tên Rate, sau đó xóa nội dung các ô trong phạm vi đó theo yêu cầu, lưu file rồi chấm lại.",
            ["P18-T02"] = "Mở sheet Exchange Rates, chọn B4:D8, đặt Number Format hiển thị tối đa 2 chữ số thập phân, lưu file rồi chấm lại.",
            ["P18-T03"] = "Mở sheet New Accounts, trong bảng tìm dòng chứa Tailspin Toys và dùng Delete Table Rows để xóa dòng đó, không xóa dữ liệu ngoài bảng, lưu file rồi chấm lại.",
            ["P18-T04"] = "Mở sheet Key Accounts cột Monthly Average, nhập hàm AVERAGE tính trung bình các tháng 1 đến 4 cho từng account, fill xuống toàn bộ cột, lưu file rồi chấm lại.",
            ["P18-T05"] = "Mở sheet Contact cột Email Address, dùng công thức ghép First Name với @woodgrovebank.com cho từng người, fill xuống toàn bộ cột, lưu file rồi chấm lại.",
            ["P18-T06"] = "Mở sheet New Accounts, chọn chart Account Balances, dùng Switch Row/Column để Opening Balance và Current Balance hiển thị trong legend, lưu file rồi chấm lại.",

            ["P20-T01"] = "Mở sheet London, chọn ô E5 và kéo fill handle/copy công thức xuống đến cuối cột của bảng, lưu file rồi chấm lại.",
            ["P20-T02"] = "Mở sheet London, chọn vùng/bảng có conditional formatting, vào Conditional Formatting > Clear Rules để xóa tất cả rule trong sheet/vùng yêu cầu, lưu file rồi chấm lại.",
            ["P20-T03"] = "Mở sheet New York City, sort bảng nhiều cấp: Country or region A-Z trước, sau đó City A-Z, lưu file rồi chấm lại.",
            ["P20-T04"] = "Mở sheet New York City ô D23, nhập hàm MAX cho cột Air Miles, lưu file rồi chấm lại.",
            ["P20-T05"] = "Mở sheet New York City, chọn dữ liệu thành phố và Air Miles, chèn Clustered Column chart, đặt chart dưới bảng với đúng vị trí/kích thước yêu cầu, lưu file rồi chấm lại.",
            ["P20-T06"] = "Mở sheet London, chọn chart Air Miles, bật Data Table và chọn tùy chọn không hiển thị legend keys, lưu file rồi chấm lại.",

            ["P22-T01"] = "Mở sheet Task, copy định dạng tiêu đề và phụ đề bằng Format Painter, sang sheet Project áp dụng cho tiêu đề và phụ đề tương ứng, lưu file rồi chấm lại.",
            ["P22-T02"] = "Mở sheet Task, chọn bảng, vào Table Design > Table Name, đặt tên bảng là Task, nhấn Enter, lưu file rồi chấm lại.",
            ["P22-T03"] = "Mở sheet Task, chọn bảng, vào Table Design > Table Style Options và bật Banded Rows để các hàng tự động tô màu xen kẽ, lưu file rồi chấm lại.",
            ["P22-T04"] = "Mở sheet Scoring Criteria ô B28, nhập công thức SUM dùng các named range Total 1, Total 2 và Total 3 thay vì tham chiếu ô trực tiếp, lưu file rồi chấm lại.",
            ["P22-T05"] = "Mở sheet Exams ô E35, nhập công thức đếm số học sinh không đạt điểm Exam 3 theo đúng điều kiện yêu cầu, lưu file rồi chấm lại.",
            ["P22-T06"] = "Mở sheet Results Distribution,Chọn biểu đồ, xóa Legend và bật Data Labels ở phía trên mỗi cột, lưu file rồi chấm lại."
        };
    }
}
