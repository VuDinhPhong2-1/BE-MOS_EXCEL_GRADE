using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Utilities;
namespace MOS.ExcelGrading.Core.Services
{
    public class XmlGradingRuleService : IXmlGradingRuleService
    {
        private const decimal StandardProjectMaxScore = 125m;
        private const int PerceptualHashThreshold = 10;
        private static readonly IReadOnlyDictionary<int, string> BuiltInExcelNumberFormats = new Dictionary<int, string>
        {
            [0] = "General",
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [14] = "m/d/yyyy",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "m/d/yyyy h:mm",
            [37] = "#,##0 ;(#,##0)",
            [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)",
            [40] = "#,##0.00;[Red](#,##0.00)",
            [44] = "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)"
        };
        private const string CommonOfficeNamespaceDeclarations =
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
            "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
            "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
            "xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
            "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " +
            "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" " +
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
            "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
            "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
            "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";
        private readonly IMongoCollection<GradingRuleSet> _ruleSets;
        private readonly ILogger<XmlGradingRuleService>? _logger;

        public XmlGradingRuleService(
            IMongoDatabase database,
            ILogger<XmlGradingRuleService>? logger = null)
        {
            _ruleSets = database.GetCollection<GradingRuleSet>("grading_rule_sets");
            _logger = logger;
        }

        public async Task<List<GradingRuleSet>> GetRuleSetsAsync(string? subject = null, bool? isActive = null)
        {
            var filter = BuildRuleSetListFilter(subject, isActive);

            return await _ruleSets.Find(filter).ToListAsync();
        }

        public async Task<List<GradingRuleSetSummary>> GetRuleSetSummariesAsync(string? subject = null, bool? isActive = null)
        {
            var projection = Builders<GradingRuleSet>.Projection
                .Include(ruleSet => ruleSet.Id)
                .Include(ruleSet => ruleSet.Subject)
                .Include(ruleSet => ruleSet.Version)
                .Include(ruleSet => ruleSet.IsActive)
                .Include("projects.projectCode")
                .Include("projects.maxScore")
                .Include("projects.tasks.taskId")
                .Include("projects.tasks.conditions.conditionId");

            var lightRuleSets = await _ruleSets
                .Find(BuildRuleSetListFilter(subject, isActive))
                .Project<GradingRuleSet>(projection)
                .ToListAsync();

            return lightRuleSets
                .Select(ruleSet => new GradingRuleSetSummary
                {
                    Id = ruleSet.Id,
                    Subject = ruleSet.Subject,
                    Version = ruleSet.Version,
                    IsActive = ruleSet.IsActive,
                    ProjectCount = ruleSet.Projects.Count,
                    TaskCount = ruleSet.Projects.Sum(project => project.Tasks.Count),
                    ConditionCount = ruleSet.Projects.Sum(project => project.Tasks.Sum(task => task.Conditions.Count)),
                    MaxScore = ruleSet.Projects.Sum(project => project.MaxScore)
                })
                .ToList();
        }

        private static FilterDefinition<GradingRuleSet> BuildRuleSetListFilter(string? subject, bool? isActive)
        {
            var filters = new List<FilterDefinition<GradingRuleSet>>();
            var normalizedSubject = NormalizeKey(subject ?? string.Empty);

            if (!string.IsNullOrWhiteSpace(normalizedSubject))
            {
                filters.Add(Builders<GradingRuleSet>.Filter.Eq(ruleSet => ruleSet.Subject, normalizedSubject));
            }

            if (isActive.HasValue)
            {
                filters.Add(Builders<GradingRuleSet>.Filter.Eq(ruleSet => ruleSet.IsActive, isActive.Value));
            }

            return filters.Count == 0
                ? Builders<GradingRuleSet>.Filter.Empty
                : Builders<GradingRuleSet>.Filter.And(filters);
        }

        public async Task<GradingRuleSet?> GetRuleSetByIdAsync(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            return await _ruleSets.Find(ruleSet => ruleSet.Id == id).FirstOrDefaultAsync();
        }

        public async Task<GradingRuleSet> CreateRuleSetAsync(GradingRuleSet ruleSet)
        {
            NormalizeRuleSetForPersistence(ruleSet);
            ValidateRuleSetShell(ruleSet);
            ruleSet.Id = string.Empty;
            await _ruleSets.InsertOneAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> UpdateRuleSetAsync(string id, GradingRuleSet ruleSet)
        {
            var existing = await GetRuleSetByIdAsync(id);
            if (existing == null)
            {
                return null;
            }

            NormalizeRuleSetForPersistence(ruleSet);
            ValidateRuleSetShell(ruleSet);
            ruleSet.Id = existing.Id;

            await _ruleSets.ReplaceOneAsync(current => current.Id == id, ruleSet);
            return ruleSet;
        }

        public async Task<bool> DeleteRuleSetAsync(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return false;
            }

            var result = await _ruleSets.DeleteOneAsync(ruleSet => ruleSet.Id == id);
            return result.DeletedCount > 0;
        }

        public async Task<GradingRuleSet?> AddProjectAsync(string ruleSetId, ProjectXmlRule project)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            if (ruleSet == null)
            {
                return null;
            }

            NormalizeProjectForPersistence(project);
            ValidateProjectShell(project);

            if (ruleSet.Projects.Any(existing => string.Equals(existing.ProjectCode, project.ProjectCode, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Project {project.ProjectCode} Ä‘Ã£ tá»“n táº¡i trong ruleset.");
            }

            ruleSet.Projects.Add(project);
            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> UpdateProjectAsync(string ruleSetId, string projectCode, ProjectXmlRule project)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            if (ruleSet == null)
            {
                return null;
            }

            NormalizeProjectForPersistence(project);
            ValidateProjectShell(project);

            var normalizedProjectCode = NormalizeKey(projectCode);
            var index = ruleSet.Projects.FindIndex(existing => string.Equals(existing.ProjectCode, normalizedProjectCode, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
            {
                return null;
            }

            project.ProjectCode = normalizedProjectCode;
            ruleSet.Projects[index] = project;
            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> DeleteProjectAsync(string ruleSetId, string projectCode)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            if (ruleSet == null)
            {
                return null;
            }

            var removed = ruleSet.Projects.RemoveAll(project => string.Equals(project.ProjectCode, NormalizeKey(projectCode), StringComparison.OrdinalIgnoreCase));
            if (removed == 0)
            {
                return null;
            }

            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> AddTaskAsync(string ruleSetId, string projectCode, TaskXmlRule task)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            var project = FindProject(ruleSet, projectCode);
            if (ruleSet == null || project == null)
            {
                return null;
            }

            NormalizeTaskForPersistence(task);
            ValidateTaskShell(task);
            EnsureSpecialConditionSupportedForSubject(task.SpecialCondition, ruleSet.Subject);

            if (project.Tasks.Any(existing => string.Equals(existing.TaskId, task.TaskId, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Task {task.TaskId} Ä‘Ã£ tá»“n táº¡i trong project {project.ProjectCode}.");
            }

            project.Tasks.Add(task);
            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> UpdateTaskAsync(string ruleSetId, string projectCode, string taskId, TaskXmlRule task)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            var project = FindProject(ruleSet, projectCode);
            if (ruleSet == null || project == null)
            {
                return null;
            }

            NormalizeTaskForPersistence(task);
            ValidateTaskShell(task);
            EnsureSpecialConditionSupportedForSubject(task.SpecialCondition, ruleSet.Subject);

            var index = project.Tasks.FindIndex(existing => string.Equals(existing.TaskId, taskId, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
            {
                return null;
            }

            task.TaskId = taskId.Trim();
            project.Tasks[index] = task;
            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> DeleteTaskAsync(string ruleSetId, string projectCode, string taskId)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            var project = FindProject(ruleSet, projectCode);
            if (ruleSet == null || project == null)
            {
                return null;
            }

            var removed = project.Tasks.RemoveAll(task => string.Equals(task.TaskId, taskId, StringComparison.OrdinalIgnoreCase));
            if (removed == 0)
            {
                return null;
            }

            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> AddConditionAsync(string ruleSetId, string projectCode, string taskId, XmlGradingCondition condition)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            var task = FindTask(FindProject(ruleSet, projectCode), taskId);
            if (ruleSet == null || task == null)
            {
                return null;
            }

            NormalizeConditionForPersistence(condition);
            ValidateConditionShell(condition);
            EnsureMinOccurrencesIsValid(condition);

            if (task.Conditions.Any(existing => string.Equals(existing.ConditionId, condition.ConditionId, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Condition {condition.ConditionId} Ä‘Ã£ tá»“n táº¡i trong task {task.TaskId}.");
            }

            task.Conditions.Add(condition);
            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> UpdateConditionAsync(string ruleSetId, string projectCode, string taskId, string conditionId, XmlGradingCondition condition)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            var task = FindTask(FindProject(ruleSet, projectCode), taskId);
            if (ruleSet == null || task == null)
            {
                return null;
            }

            NormalizeConditionForPersistence(condition);
            ValidateConditionShell(condition);
            EnsureMinOccurrencesIsValid(condition);

            var index = task.Conditions.FindIndex(existing => string.Equals(existing.ConditionId, conditionId, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
            {
                return null;
            }

            condition.ConditionId = conditionId.Trim();
            task.Conditions[index] = condition;
            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> DeleteConditionAsync(string ruleSetId, string projectCode, string taskId, string conditionId)
        {
            var ruleSet = await GetRuleSetByIdAsync(ruleSetId);
            var task = FindTask(FindProject(ruleSet, projectCode), taskId);
            if (ruleSet == null || task == null)
            {
                return null;
            }

            var removed = task.Conditions.RemoveAll(condition => string.Equals(condition.ConditionId, conditionId, StringComparison.OrdinalIgnoreCase));
            if (removed == 0)
            {
                return null;
            }

            await ReplaceRuleSetAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingRuleSet?> GetActiveRuleSetAsync(string subject, string projectCode)
        {
            var normalizedSubject = NormalizeKey(subject);
            var normalizedProjectCode = NormalizeKey(projectCode);

            if (string.IsNullOrWhiteSpace(normalizedSubject) || string.IsNullOrWhiteSpace(normalizedProjectCode))
            {
                return null;
            }

            var filter = Builders<GradingRuleSet>.Filter.And(
                Builders<GradingRuleSet>.Filter.Eq(ruleSet => ruleSet.Subject, normalizedSubject),
                Builders<GradingRuleSet>.Filter.Eq(ruleSet => ruleSet.IsActive, true),
                Builders<GradingRuleSet>.Filter.ElemMatch(
                    ruleSet => ruleSet.Projects,
                    project => project.ProjectCode == normalizedProjectCode));

            var projection = Builders<GradingRuleSet>.Projection
                .Include(ruleSet => ruleSet.Id)
                .Include(ruleSet => ruleSet.Subject)
                .Include(ruleSet => ruleSet.Version)
                .Include(ruleSet => ruleSet.IsActive)
                .ElemMatch(ruleSet => ruleSet.Projects, project => project.ProjectCode == normalizedProjectCode);

            return await _ruleSets
                .Find(filter)
                .Project<GradingRuleSet>(projection)
                .FirstOrDefaultAsync();
        }

        public Task<XmlRuleValidationResult> ValidateRuleSetAsync(GradingRuleSet ruleSet)
        {
            return Task.FromResult(ValidateRuleSet(ruleSet));
        }

        public async Task<GradingRuleSet> SeedProject22Task1RuleSetAsync()
        {
            var ruleSet = BuildProject22Task1RuleSet();
            var validation = ValidateRuleSet(ruleSet);
            if (!validation.IsValid)
            {
                throw new InvalidOperationException($"Seed XML grading ruleset is invalid: {string.Join("; ", validation.Errors)}");
            }

            var filter = Builders<GradingRuleSet>.Filter.And(
                Builders<GradingRuleSet>.Filter.Eq(existing => existing.Subject, ruleSet.Subject),
                Builders<GradingRuleSet>.Filter.Eq(existing => existing.Version, ruleSet.Version),
                Builders<GradingRuleSet>.Filter.ElemMatch(
                    existing => existing.Projects,
                    project => project.ProjectCode == "project22"));

            var existingRuleSet = await _ruleSets.Find(filter).FirstOrDefaultAsync();
            if (existingRuleSet != null)
            {
                ruleSet.Id = existingRuleSet.Id;
                await _ruleSets.ReplaceOneAsync(
                    Builders<GradingRuleSet>.Filter.Eq(existing => existing.Id, existingRuleSet.Id),
                    ruleSet);
                return ruleSet;
            }

            await _ruleSets.InsertOneAsync(ruleSet);
            return ruleSet;
        }

        public async Task<GradingResult> GradeAsync(Stream studentFile, string subject, string projectCode)
        {
            var totalStopwatch = Stopwatch.StartNew();
            var phaseStopwatch = Stopwatch.StartNew();

            var ruleSet = await GetActiveRuleSetAsync(subject, projectCode)
                ?? throw new InvalidOperationException($"KhÃ´ng tÃ¬m tháº¥y XML grading ruleset active cho {subject}/{projectCode}.");

            var rulesetMs = phaseStopwatch.ElapsedMilliseconds;

            var projectRule = ruleSet.Projects.FirstOrDefault(project =>
                string.Equals(project.ProjectCode, NormalizeKey(projectCode), StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"KhÃ´ng tÃ¬m tháº¥y project rule cho {projectCode}.");

            phaseStopwatch.Restart();
            var requiredParts = CollectRequiredOfficeParts(projectRule);
            var package = ReadOfficePackage(studentFile, requiredParts);
            var packageMs = phaseStopwatch.ElapsedMilliseconds;

            var result = new GradingResult
            {
                ProjectId = projectRule.ProjectCode,
                ProjectName = string.IsNullOrWhiteSpace(projectRule.ProjectName) ? projectRule.ProjectCode : projectRule.ProjectName,
                MaxScore = StandardProjectMaxScore
            };

            var evaluationCache = new XmlEvaluationCache();
            long specialMs = 0;
            long conditionsMs = 0;

            foreach (var taskRule in projectRule.Tasks)
            {
                var taskResult = new TaskResult
                {
                    TaskId = taskRule.TaskId,
                    TaskName = taskRule.TaskName,
                    MaxScore = taskRule.MaxScore
                };

                // ===== SPECIAL CONDITION (cáº¥p Task) =====
                // Hoáº¡t Ä‘á»™ng nhÆ° má»™t "gate" + cÃ³ Ä‘iá»ƒm riÃªng: náº¿u Task cÃ³ specialCondition
                // mÃ  nÃ³ FAIL, toÃ n bá»™ Task = 0 Ä‘iá»ƒm báº¥t ká»ƒ conditions XML thÆ°á»ng cÃ³ Ä‘Ãºng
                // hay khÃ´ng. Náº¿u PASS, cá»™ng thÃªm SpecialCondition.Score vÃ o Ä‘iá»ƒm Task
                // (Ä‘á»™c láº­p vá»›i Ä‘iá»ƒm cÃ¡c Conditions XML, khÃ´ng cÃ²n "Äƒn trá»n" MaxScore).
                var hasSpecialCondition = taskRule.SpecialCondition != null
                    && !string.IsNullOrWhiteSpace(taskRule.SpecialCondition.Type);

                var specialConditionPassed = true;

                if (hasSpecialCondition)
                {
                    phaseStopwatch.Restart();
                    var specialResult = EvaluateTaskSpecialCondition(taskRule.SpecialCondition!, package, ruleSet.Subject);
                    specialMs += phaseStopwatch.ElapsedMilliseconds;
                    specialConditionPassed = specialResult.IsPassed;

                    if (specialResult.IsPassed)
                    {
                        var successDetail = taskRule.SpecialCondition!.Feedback?.SuccessDetail?.Trim() ?? string.Empty;
                        taskResult.Details.Add(string.IsNullOrWhiteSpace(successDetail)
                            ? $"[SpecialCondition:{taskRule.SpecialCondition!.Type}] {specialResult.Message}"
                            : successDetail);

                        // Cá»™ng Ä‘iá»ƒm riÃªng cá»§a specialCondition do ngÆ°á»i táº¡o ruleset cáº¥u hÃ¬nh.
                        // Cho phÃ©p Task chá»‰ dÃ¹ng specialCondition (0 condition XML) hoáº·c
                        // káº¿t há»£p cáº£ hai, miá»…n tá»•ng = task.maxScore (Ä‘Æ°á»£c validate á»Ÿ ValidateRuleSet).
                        taskResult.Score += taskRule.SpecialCondition!.Score;
                    }
                    else
                    {
                        var fallbackMessage = $"[SpecialCondition:{taskRule.SpecialCondition!.Type}] {specialResult.Message}";
                        var message = taskRule.SpecialCondition!.Feedback?.ErrorMessage?.Trim() ?? string.Empty;
                        var fixAction = taskRule.SpecialCondition!.Feedback?.FixAction?.Trim() ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(message))
                        {
                            message = fallbackMessage;
                        }
                        else
                        {
                            taskResult.Details.Add(fallbackMessage);
                        }

                        taskResult.Errors.Add(message);
                        taskResult.DisplayIssues.Add(new TaskDisplayIssue
                        {
                            Heading = string.IsNullOrWhiteSpace(taskRule.TaskName) ? taskRule.TaskId : taskRule.TaskName,
                            Message = message,
                            FixAction = fixAction,
                        });

                        if (!string.IsNullOrWhiteSpace(fixAction))
                        {
                            taskResult.FixActions.Add(fixAction);
                        }
                    }
                }

                // ===== NORMAL XML CONDITIONS =====
                phaseStopwatch.Restart();
                foreach (var condition in taskRule.Conditions)
                {
                    var conditionResult = EvaluateCondition(condition, package, evaluationCache);
                    if (conditionResult.IsPassed)
                    {
                        taskResult.Score += conditionResult.ScoreAwarded;
                        if (!string.IsNullOrWhiteSpace(conditionResult.Feedback.SuccessDetail))
                        {
                            taskResult.Details.Add(conditionResult.Feedback.SuccessDetail.Trim());
                        }
                    }
                    else
                    {
                        var errorMessage = conditionResult.Feedback.ErrorMessage?.Trim() ?? string.Empty;
                        var fixAction = conditionResult.Feedback.FixAction?.Trim() ?? string.Empty;

                        if (!string.IsNullOrWhiteSpace(errorMessage))
                        {
                            taskResult.Errors.Add(errorMessage);
                            taskResult.DisplayIssues.Add(new TaskDisplayIssue
                            {
                                Heading = string.IsNullOrWhiteSpace(taskRule.TaskName) ? taskRule.TaskId : taskRule.TaskName,
                                Message = errorMessage,
                                FixAction = fixAction,
                            });
                        }

                        if (!string.IsNullOrWhiteSpace(fixAction))
                        {
                            taskResult.FixActions.Add(fixAction);
                        }
                    }

                    taskResult.Details.Add(
                        $"XML condition {conditionResult.ConditionId}: sourceFile={conditionResult.SourceFile}, compareMode={conditionResult.CompareMode}, matchPolicy={conditionResult.MatchPolicy}, matched={conditionResult.MatchedExpectedValues.Count}, missing={conditionResult.MissingExpectedValues.Count}.");

                    if (!conditionResult.IsPassed && condition.StopTaskIfFailed)
                    {
                        break;
                    }
                }
                // Special condition FAIL -> zero toÃ n bá»™ Task, ká»ƒ cáº£ khi cÃ³ conditions XML Ä‘Ã£ Ä‘áº¡t Ä‘iá»ƒm.
                conditionsMs += phaseStopwatch.ElapsedMilliseconds;
                if (hasSpecialCondition && !specialConditionPassed)
                {
                    taskResult.Score = 0m;
                }

                if (taskResult.Score > taskResult.MaxScore)
                {
                    taskResult.Score = taskResult.MaxScore;
                }

                result.TaskResults.Add(taskResult);
            }

            ApplyProjectScoringModel(result);
            totalStopwatch.Stop();
            _logger?.LogInformation(
                "XML grading timing subject={Subject} project={ProjectCode} tasks={TaskCount} rulesetMs={RuleSetMs} packageMs={PackageMs} specialMs={SpecialMs} conditionsMs={ConditionsMs} totalMs={TotalMs} xmlParts={XmlPartCount} binaryParts={BinaryPartCount}",
                NormalizeKey(subject),
                NormalizeKey(projectCode),
                projectRule.Tasks.Count,
                rulesetMs,
                packageMs,
                specialMs,
                conditionsMs,
                totalStopwatch.ElapsedMilliseconds,
                package.XmlParts.Count,
                package.BinaryParts.Count);
            return result;
        }

        private async Task ReplaceRuleSetAsync(GradingRuleSet ruleSet)
        {
            await _ruleSets.ReplaceOneAsync(current => current.Id == ruleSet.Id, ruleSet);
        }

        private static ProjectXmlRule? FindProject(GradingRuleSet? ruleSet, string projectCode)
        {
            var normalizedProjectCode = NormalizeKey(projectCode);
            return ruleSet?.Projects.FirstOrDefault(project =>
                string.Equals(project.ProjectCode, normalizedProjectCode, StringComparison.OrdinalIgnoreCase));
        }

        private static TaskXmlRule? FindTask(ProjectXmlRule? project, string taskId)
        {
            return project?.Tasks.FirstOrDefault(task =>
                string.Equals(task.TaskId, taskId, StringComparison.OrdinalIgnoreCase));
        }

        private static void NormalizeRuleSetForPersistence(GradingRuleSet ruleSet)
        {
            ruleSet.Subject = NormalizeKey(ruleSet.Subject);
            ruleSet.Version = string.IsNullOrWhiteSpace(ruleSet.Version) ? "v1" : ruleSet.Version.Trim();
            ruleSet.Projects ??= new List<ProjectXmlRule>();
            foreach (var project in ruleSet.Projects)
            {
                NormalizeProjectForPersistence(project);
            }
        }

        private static void NormalizeProjectForPersistence(ProjectXmlRule project)
        {
            project.ProjectCode = NormalizeKey(project.ProjectCode);
            project.ProjectName = project.ProjectName?.Trim() ?? string.Empty;
            project.Tasks ??= new List<TaskXmlRule>();
            foreach (var task in project.Tasks)
            {
                NormalizeTaskForPersistence(task);
            }
        }

        private static void NormalizeTaskForPersistence(TaskXmlRule task)
        {
            task.TaskId = task.TaskId?.Trim() ?? string.Empty;
            task.TaskName = task.TaskName?.Trim() ?? string.Empty;
            task.Conditions ??= new List<XmlGradingCondition>();
            foreach (var condition in task.Conditions)
            {
                NormalizeConditionForPersistence(condition);
            }

            if (task.SpecialCondition != null)
            {
                task.SpecialCondition.Type = task.SpecialCondition.Type?.Trim() ?? string.Empty;

                // Náº¿u FE gá»­i type rá»—ng (ngÆ°á»i dÃ¹ng chá»n "KhÃ´ng sá»­ dá»¥ng"), coi nhÆ° khÃ´ng cÃ³ special condition.
                if (string.IsNullOrWhiteSpace(task.SpecialCondition.Type))
                {
                    task.SpecialCondition = null;
                }
                else
                {
                    if (task.SpecialCondition.Score < 0)
                    {
                        task.SpecialCondition.Score = 0m;
                    }

                    task.SpecialCondition.Feedback ??= new ConditionFeedback();
                    task.SpecialCondition.Feedback.SuccessDetail = task.SpecialCondition.Feedback.SuccessDetail?.Trim() ?? string.Empty;
                    task.SpecialCondition.Feedback.ErrorMessage = task.SpecialCondition.Feedback.ErrorMessage?.Trim() ?? string.Empty;
                    task.SpecialCondition.Feedback.FixAction = task.SpecialCondition.Feedback.FixAction?.Trim() ?? string.Empty;

                    // ===== Normalize cho type = pictureBullet =====
                    if (task.SpecialCondition.Config != null)
                    {
                        task.SpecialCondition.Config.AssetId = string.IsNullOrWhiteSpace(task.SpecialCondition.Config.AssetId)
                            ? null
                            : task.SpecialCondition.Config.AssetId.Trim();

                        task.SpecialCondition.Config.ImageHash = string.IsNullOrWhiteSpace(task.SpecialCondition.Config.ImageHash)
                            ? null
                            : ImageHashUtility.NormalizeHash(task.SpecialCondition.Config.ImageHash);

                        // Má»šI: dHash luÃ´n lÃ  hex 16 kÃ½ tá»± cá»‘ Ä‘á»‹nh, chá»‰ cáº§n trim + lowercase,
                        // khÃ´ng cáº§n NormalizeHash (hÃ m Ä‘Ã³ dÃ nh riÃªng cho SHA-256).
                        task.SpecialCondition.Config.PerceptualHash = string.IsNullOrWhiteSpace(task.SpecialCondition.Config.PerceptualHash)
                            ? null
                            : task.SpecialCondition.Config.PerceptualHash.Trim().ToLowerInvariant();
                    }

                    // ===== Normalize cho type = insertedImage =====
                    // TÃ¡ch riÃªng khá»‘i config nÃ y vá»›i pictureBullet á»Ÿ trÃªn vÃ¬ 2 loáº¡i
                    // special condition dÃ¹ng 2 property config khÃ¡c nhau
                    // (Config vs ImageInsertConfig), khÃ´ng lá»“ng chung 1 object.
                    if (task.SpecialCondition.ImageInsertConfig != null)
                    {
                        var imageInsertConfig = task.SpecialCondition.ImageInsertConfig;

                        imageInsertConfig.SourceFile = string.IsNullOrWhiteSpace(imageInsertConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(imageInsertConfig.SourceFile);

                        imageInsertConfig.RelsFile = string.IsNullOrWhiteSpace(imageInsertConfig.RelsFile)
                            ? "word/_rels/document.xml.rels"
                            : NormalizeSourceFile(imageInsertConfig.RelsFile);

                        imageInsertConfig.AssetId = string.IsNullOrWhiteSpace(imageInsertConfig.AssetId)
                            ? null
                            : imageInsertConfig.AssetId.Trim();

                        imageInsertConfig.ImageHash = string.IsNullOrWhiteSpace(imageInsertConfig.ImageHash)
                            ? null
                            : ImageHashUtility.NormalizeHash(imageInsertConfig.ImageHash);

                        // Má»šI
                        imageInsertConfig.PerceptualHash = string.IsNullOrWhiteSpace(imageInsertConfig.PerceptualHash)
                            ? null
                            : imageInsertConfig.PerceptualHash.Trim().ToLowerInvariant();

                        // wrapType lÃ  optional: rá»—ng nghÄ©a lÃ  khÃ´ng cáº§n kiá»ƒm tra
                        // cháº¿ Ä‘á»™ ngáº¯t dÃ²ng, chá»‰ kiá»ƒm tra Ä‘Ãºng áº£nh. Trim + Ä‘á»ƒ null
                        // náº¿u rá»—ng Ä‘á»ƒ trÃ¡nh lÆ°u chuá»—i khoáº£ng tráº¯ng vÃ o DB.
                        imageInsertConfig.WrapType = string.IsNullOrWhiteSpace(imageInsertConfig.WrapType)
                            ? null
                            : imageInsertConfig.WrapType.Trim();

                        if (imageInsertConfig.PositionConfig != null)
                        {
                            imageInsertConfig.PositionConfig.AfterText = string.IsNullOrWhiteSpace(imageInsertConfig.PositionConfig.AfterText)
                                ? null
                                : NormalizePlainText(imageInsertConfig.PositionConfig.AfterText);

                            imageInsertConfig.PositionConfig.BeforeText = string.IsNullOrWhiteSpace(imageInsertConfig.PositionConfig.BeforeText)
                                ? null
                                : NormalizePlainText(imageInsertConfig.PositionConfig.BeforeText);

                            imageInsertConfig.PositionConfig.RequireBetween ??= true;
                            imageInsertConfig.PositionConfig.CaseSensitive ??= false;
                        }

                        if (imageInsertConfig.SizeConfig != null)
                        {
                            if (imageInsertConfig.SizeConfig.ExpectedWidthEmu <= 0)
                            {
                                imageInsertConfig.SizeConfig.ExpectedWidthEmu = null;
                            }

                            if (imageInsertConfig.SizeConfig.ExpectedHeightEmu <= 0)
                            {
                                imageInsertConfig.SizeConfig.ExpectedHeightEmu = null;
                            }

                            if (imageInsertConfig.SizeConfig.ToleranceEmu < 0)
                            {
                                imageInsertConfig.SizeConfig.ToleranceEmu = 0;
                            }
                        }
                    }

                    if (task.SpecialCondition.ConvertTableToTextConfig != null)
                    {
                        var convertConfig = task.SpecialCondition.ConvertTableToTextConfig;

                        convertConfig.SourceFile = string.IsNullOrWhiteSpace(convertConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(convertConfig.SourceFile);

                        convertConfig.AnchorText = string.IsNullOrWhiteSpace(convertConfig.AnchorText)
                            ? null
                            : NormalizePlainText(convertConfig.AnchorText);

                        convertConfig.ExpectedRows = convertConfig.ExpectedRows?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizePlainText)
                            .ToList() ?? new List<string>();

                        if (convertConfig.MinRows <= 0)
                        {
                            convertConfig.MinRows = null;
                        }

                        if (convertConfig.MinTabsPerRow <= 0)
                        {
                            convertConfig.MinTabsPerRow = null;
                        }

                        convertConfig.RequireNoTables ??= true;
                    }

                    if (task.SpecialCondition.HyperlinkConfig != null)
                    {
                        var hyperlinkConfig = task.SpecialCondition.HyperlinkConfig;

                        hyperlinkConfig.SourceFile = string.IsNullOrWhiteSpace(hyperlinkConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(hyperlinkConfig.SourceFile);

                        hyperlinkConfig.RelsFile = string.IsNullOrWhiteSpace(hyperlinkConfig.RelsFile)
                            ? "word/_rels/document.xml.rels"
                            : NormalizeSourceFile(hyperlinkConfig.RelsFile);

                        hyperlinkConfig.DisplayText = string.IsNullOrWhiteSpace(hyperlinkConfig.DisplayText)
                            ? null
                            : NormalizePlainText(hyperlinkConfig.DisplayText);

                        hyperlinkConfig.AnchorTextBefore = string.IsNullOrWhiteSpace(hyperlinkConfig.AnchorTextBefore)
                            ? null
                            : NormalizePlainText(hyperlinkConfig.AnchorTextBefore);

                        hyperlinkConfig.Url = string.IsNullOrWhiteSpace(hyperlinkConfig.Url)
                            ? null
                            : hyperlinkConfig.Url.Trim();

                        hyperlinkConfig.CaseSensitiveText ??= false;
                    }

                    if (task.SpecialCondition.SectionBreakBeforeTextConfig != null)
                    {
                        var sectionConfig = task.SpecialCondition.SectionBreakBeforeTextConfig;

                        sectionConfig.SourceFile = string.IsNullOrWhiteSpace(sectionConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(sectionConfig.SourceFile);

                        sectionConfig.TargetText = string.IsNullOrWhiteSpace(sectionConfig.TargetText)
                            ? null
                            : NormalizePlainText(sectionConfig.TargetText);

                        sectionConfig.BreakType = string.IsNullOrWhiteSpace(sectionConfig.BreakType)
                            ? "continuous"
                            : sectionConfig.BreakType.Trim();

                        if (sectionConfig.TargetOccurrence <= 0)
                        {
                            sectionConfig.TargetOccurrence = 1;
                        }

                        sectionConfig.RequireImmediateBefore ??= true;
                        sectionConfig.AllowSameParagraphSectPr ??= true;
                    }

                    if (task.SpecialCondition.PictureStyleConfig != null)
                    {
                        var pictureStyleConfig = task.SpecialCondition.PictureStyleConfig;

                        pictureStyleConfig.SourceFile = string.IsNullOrWhiteSpace(pictureStyleConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(pictureStyleConfig.SourceFile);

                        pictureStyleConfig.RelsFile = string.IsNullOrWhiteSpace(pictureStyleConfig.RelsFile)
                            ? "word/_rels/document.xml.rels"
                            : NormalizeSourceFile(pictureStyleConfig.RelsFile);

                        pictureStyleConfig.AssetId = string.IsNullOrWhiteSpace(pictureStyleConfig.AssetId)
                            ? null
                            : pictureStyleConfig.AssetId.Trim();

                        pictureStyleConfig.ImageHash = string.IsNullOrWhiteSpace(pictureStyleConfig.ImageHash)
                            ? null
                            : ImageHashUtility.NormalizeHash(pictureStyleConfig.ImageHash);

                        pictureStyleConfig.PerceptualHash = string.IsNullOrWhiteSpace(pictureStyleConfig.PerceptualHash)
                            ? null
                            : pictureStyleConfig.PerceptualHash.Trim().ToLowerInvariant();

                        if (pictureStyleConfig.TargetImageIndex <= 0)
                        {
                            pictureStyleConfig.TargetImageIndex = 1;
                        }

                        pictureStyleConfig.StylePreset = string.IsNullOrWhiteSpace(pictureStyleConfig.StylePreset)
                            ? "simpleFrameBlack"
                            : pictureStyleConfig.StylePreset.Trim();

                        pictureStyleConfig.RequiredLineColor = NormalizeHexColor(pictureStyleConfig.RequiredLineColor);

                        if (pictureStyleConfig.MinLineWidth <= 0)
                        {
                            pictureStyleConfig.MinLineWidth = null;
                        }

                        pictureStyleConfig.PresetGeometry = string.IsNullOrWhiteSpace(pictureStyleConfig.PresetGeometry)
                            ? null
                            : pictureStyleConfig.PresetGeometry.Trim();
                    }

                    if (task.SpecialCondition.TextBoxContainsTextConfig != null)
                    {
                        var textBoxConfig = task.SpecialCondition.TextBoxContainsTextConfig;

                        textBoxConfig.SourceFile = string.IsNullOrWhiteSpace(textBoxConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(textBoxConfig.SourceFile);

                        textBoxConfig.ExpectedText = string.IsNullOrWhiteSpace(textBoxConfig.ExpectedText)
                            ? null
                            : NormalizePlainText(textBoxConfig.ExpectedText);

                        textBoxConfig.MatchMode = string.IsNullOrWhiteSpace(textBoxConfig.MatchMode)
                            ? "exact"
                            : textBoxConfig.MatchMode.Trim();

                        textBoxConfig.CaseSensitive ??= false;

                        if (textBoxConfig.TargetOccurrence <= 0)
                        {
                            textBoxConfig.TargetOccurrence = 1;
                        }

                        textBoxConfig.RequireDefaultPaste ??= true;
                        textBoxConfig.RequireRemovedFromBody ??= true;
                        textBoxConfig.ForbiddenTextColors = textBoxConfig.ForbiddenTextColors?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizeHexColor)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(value => value!)
                            .ToList() ?? new List<string>();

                        textBoxConfig.ForbiddenRunProperties = textBoxConfig.ForbiddenRunProperties?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(value => value.Trim())
                            .ToList() ?? new List<string>();
                    }

                    if (task.SpecialCondition.PageMarginsConfig != null)
                    {
                        var marginsConfig = task.SpecialCondition.PageMarginsConfig;

                        marginsConfig.SourceFile = string.IsNullOrWhiteSpace(marginsConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(marginsConfig.SourceFile);

                        marginsConfig.RequireAllSections ??= true;
                    }

                    if (task.SpecialCondition.DocumentStyleSetConfig != null)
                    {
                        var styleSetConfig = task.SpecialCondition.DocumentStyleSetConfig;

                        styleSetConfig.SourceFile = string.IsNullOrWhiteSpace(styleSetConfig.SourceFile)
                            ? "word/styles.xml"
                            : NormalizeSourceFile(styleSetConfig.SourceFile);

                        styleSetConfig.StyleSetName = string.IsNullOrWhiteSpace(styleSetConfig.StyleSetName)
                            ? null
                            : styleSetConfig.StyleSetName.Trim();

                        styleSetConfig.ExpectedFragments = styleSetConfig.ExpectedFragments?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(value => value.Trim())
                            .ToList() ?? new List<string>();

                        styleSetConfig.IgnoreAttributes = styleSetConfig.IgnoreAttributes?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(value => value.Trim())
                            .ToList() ?? new List<string>();

                        styleSetConfig.MatchPolicy = string.IsNullOrWhiteSpace(styleSetConfig.MatchPolicy)
                            ? XmlGradingMatchPolicies.All
                            : styleSetConfig.MatchPolicy.Trim();
                    }

                    if (task.SpecialCondition.PageBorderConfig != null)
                    {
                        var pageBorderConfig = task.SpecialCondition.PageBorderConfig;

                        pageBorderConfig.SourceFile = string.IsNullOrWhiteSpace(pageBorderConfig.SourceFile)
                            ? "word/document.xml"
                            : NormalizeSourceFile(pageBorderConfig.SourceFile);

                        pageBorderConfig.RequiredStyle = string.IsNullOrWhiteSpace(pageBorderConfig.RequiredStyle)
                            ? "single"
                            : pageBorderConfig.RequiredStyle.Trim();

                        if (pageBorderConfig.RequiredWidth <= 0)
                        {
                            pageBorderConfig.RequiredWidth = null;
                        }

                        if (pageBorderConfig.MinWidth <= 0)
                        {
                            pageBorderConfig.MinWidth = null;
                        }

                        pageBorderConfig.RequiredColor = NormalizeBorderColor(pageBorderConfig.RequiredColor);
                        pageBorderConfig.AllowedColors = pageBorderConfig.AllowedColors?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizeBorderColor)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(value => value!)
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();

                        if (!string.IsNullOrWhiteSpace(pageBorderConfig.RequiredColor)
                            && !pageBorderConfig.AllowedColors.Contains(pageBorderConfig.RequiredColor, StringComparer.OrdinalIgnoreCase))
                        {
                            pageBorderConfig.AllowedColors.Insert(0, pageBorderConfig.RequiredColor);
                        }

                        pageBorderConfig.RequireBox ??= true;
                        pageBorderConfig.RequireAllSections ??= true;
                    }

                    if (task.SpecialCondition.ExcelTableNameConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelTableNameConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.ExpectedName = string.IsNullOrWhiteSpace(config.ExpectedName) ? null : config.ExpectedName.Trim();
                        config.OriginalName = string.IsNullOrWhiteSpace(config.OriginalName) ? null : config.OriginalName.Trim();
                        config.RequireOriginalNameAbsent ??= true;
                    }

                    if (task.SpecialCondition.ExcelWorksheetPageSetupConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelWorksheetPageSetupConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.Orientation = string.IsNullOrWhiteSpace(config.Orientation) ? "landscape" : config.Orientation.Trim();
                    }

                    if (task.SpecialCondition.ExcelClearCellFormattingConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelClearCellFormattingConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.Range = string.IsNullOrWhiteSpace(config.Range) ? null : config.Range.Trim().ToUpperInvariant();
                        if (config.DefaultStyleId < 0)
                        {
                            config.DefaultStyleId = 0;
                        }
                    }

                    if (task.SpecialCondition.ExcelDataModelImportConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelDataModelImportConfig;
                        config.SourceFileName = string.IsNullOrWhiteSpace(config.SourceFileName) ? null : config.SourceFileName.Trim();
                        config.ExpectedWorksheetName = string.IsNullOrWhiteSpace(config.ExpectedWorksheetName) ? null : config.ExpectedWorksheetName.Trim();
                        config.ExpectedConnectionName = string.IsNullOrWhiteSpace(config.ExpectedConnectionName) ? null : config.ExpectedConnectionName.Trim();
                        config.RequireConnection ??= true;
                        config.RequireDataModel ??= true;
                        config.RequireImportedWorksheet ??= true;
                        config.RequireQueryTable ??= true;
                    }

                    if (task.SpecialCondition.ExcelCompatibilityReportConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelCompatibilityReportConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.ExpectedTexts = config.ExpectedTexts?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizePlainText)
                            .ToList() ?? new List<string>();
                        config.RequireNewWorksheet ??= true;
                    }

                    if (task.SpecialCondition.ExcelMergedRangeConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelMergedRangeConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.Range = string.IsNullOrWhiteSpace(config.Range) ? null : NormalizeExcelRangeAddress(config.Range);
                        config.RequireNoHorizontalCenter ??= false;
                    }

                    if (task.SpecialCondition.ExcelCellHyperlinkConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelCellHyperlinkConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.Cell = string.IsNullOrWhiteSpace(config.Cell) ? null : NormalizeExcelCellAddress(config.Cell);
                        config.Location = string.IsNullOrWhiteSpace(config.Location) ? null : config.Location.Trim();
                        config.Target = string.IsNullOrWhiteSpace(config.Target) ? null : config.Target.Trim();
                        config.Display = string.IsNullOrWhiteSpace(config.Display) ? null : config.Display.Trim();
                    }

                    if (task.SpecialCondition.ExcelIconSetConditionalFormattingConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelIconSetConditionalFormattingConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.Range = string.IsNullOrWhiteSpace(config.Range) ? null : NormalizeExcelRangeAddress(config.Range);
                        config.IconSet = string.IsNullOrWhiteSpace(config.IconSet) ? "3Flags" : config.IconSet.Trim();
                    }

                    if (task.SpecialCondition.ExcelChartDataRangeConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelChartDataRangeConfig;
                        config.ChartSourceFile = string.IsNullOrWhiteSpace(config.ChartSourceFile) ? null : NormalizeSourceFile(config.ChartSourceFile);
                        config.ExpectedCategoryRange = string.IsNullOrWhiteSpace(config.ExpectedCategoryRange) ? null : NormalizeExcelFormulaReference(config.ExpectedCategoryRange);
                        config.ExpectedValueRange = string.IsNullOrWhiteSpace(config.ExpectedValueRange) ? null : NormalizeExcelFormulaReference(config.ExpectedValueRange);
                        config.ExpectedValueRanges = config.ExpectedValueRanges?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizeExcelFormulaReference)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();
                        if (!string.IsNullOrWhiteSpace(config.ExpectedValueRange)
                            && !config.ExpectedValueRanges.Contains(config.ExpectedValueRange, StringComparer.OrdinalIgnoreCase))
                        {
                            config.ExpectedValueRanges.Insert(0, config.ExpectedValueRange);
                        }
                        config.ExpectedSeriesNames = config.ExpectedSeriesNames?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizePlainText)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();
                        config.ExpectedCategoryText = string.IsNullOrWhiteSpace(config.ExpectedCategoryText) ? null : NormalizePlainText(config.ExpectedCategoryText);
                        if (config.ExpectedPointCount <= 0)
                        {
                            config.ExpectedPointCount = null;
                        }
                        config.RequireNoExtraSeries ??= true;
                    }

                    if (task.SpecialCondition.ExcelChartStyleConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelChartStyleConfig;
                        config.ChartSourceFile = string.IsNullOrWhiteSpace(config.ChartSourceFile) ? null : NormalizeSourceFile(config.ChartSourceFile);
                        config.StyleSourceFile = string.IsNullOrWhiteSpace(config.StyleSourceFile) ? null : NormalizeSourceFile(config.StyleSourceFile);
                        if (config.StyleId <= 0)
                        {
                            config.StyleId = null;
                        }
                    }

                    if (task.SpecialCondition.ExcelTextReplacementConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelTextReplacementConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.OldText = string.IsNullOrWhiteSpace(config.OldText) ? null : NormalizePlainText(config.OldText);
                        config.NewText = string.IsNullOrWhiteSpace(config.NewText) ? null : NormalizePlainText(config.NewText);
                        if (config.MinNewTextOccurrences <= 0)
                        {
                            config.MinNewTextOccurrences = 1;
                        }
                        config.RequireOldTextAbsent ??= true;
                        config.MatchWholeWord ??= true;
                    }

                    if (task.SpecialCondition.ExcelPrintTitlesConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelPrintTitlesConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.ExpectedRows = string.IsNullOrWhiteSpace(config.ExpectedRows) ? null : NormalizeExcelPrintRows(config.ExpectedRows);
                    }

                    if (task.SpecialCondition.ExcelNumberFormatConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelNumberFormatConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.Range = string.IsNullOrWhiteSpace(config.Range) ? null : NormalizeExcelRangeOrColumnRange(config.Range);
                        config.AllowedNumberFormatIds = config.AllowedNumberFormatIds?
                            .Where(value => value > 0)
                            .Distinct()
                            .ToList() ?? new List<int> { 1, 2, 3, 4 };
                        config.Category = string.IsNullOrWhiteSpace(config.Category) ? null : config.Category.Trim().ToLowerInvariant();
                        if (config.DecimalPlaces < 0)
                        {
                            config.DecimalPlaces = null;
                        }
                        config.Symbol = string.IsNullOrWhiteSpace(config.Symbol) ? null : config.Symbol.Trim();
                        config.RequireEveryNumericCell ??= true;
                    }

                    if (task.SpecialCondition.ExcelChartLegendConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelChartLegendConfig;
                        config.ChartSourceFile = string.IsNullOrWhiteSpace(config.ChartSourceFile) ? null : NormalizeSourceFile(config.ChartSourceFile);
                        config.Position = string.IsNullOrWhiteSpace(config.Position) ? "t" : config.Position.Trim();
                    }

                    if (task.SpecialCondition.ExcelDefinedNameConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelDefinedNameConfig;
                        config.Name = string.IsNullOrWhiteSpace(config.Name) ? null : config.Name.Trim();
                        config.ExpectedRanges = config.ExpectedRanges?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizeExcelFormulaReference)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();
                        config.RequireExactRanges ??= true;
                    }

                    if (task.SpecialCondition.ExcelFormulaReferencesConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelFormulaReferencesConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.Cell = string.IsNullOrWhiteSpace(config.Cell) ? null : NormalizeExcelCellAddress(config.Cell);
                        config.RequiredReferences = config.RequiredReferences?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(value => value.Trim())
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();
                        config.RequiredFunctions = config.RequiredFunctions?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(value => value.Trim().TrimEnd('(').ToUpperInvariant())
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();
                        config.RequiredFormulaFragments = config.RequiredFormulaFragments?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizeExcelFormulaFragment)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();
                        config.ExpectedFormula = string.IsNullOrWhiteSpace(config.ExpectedFormula) ? null : NormalizeExcelFormulaText(config.ExpectedFormula);
                        config.ExpectedValue = string.IsNullOrWhiteSpace(config.ExpectedValue) ? null : NormalizePlainText(config.ExpectedValue);
                        config.RequireOnlyDefinedNameReferences ??= true;
                    }

                    if (task.SpecialCondition.ExcelNoConditionalFormattingConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelNoConditionalFormattingConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.RequireAllWorksheets ??= false;
                    }

                    if (task.SpecialCondition.ExcelTextRotationConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelTextRotationConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.ExpectedTexts = config.ExpectedTexts?
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Select(NormalizePlainText)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList() ?? new List<string>();
                        config.AllowedTextRotationValues = config.AllowedTextRotationValues?
                            .Where(value => value is >= 0 and <= 180)
                            .Distinct()
                            .ToList() ?? new List<int> { 45 };
                        if (config.AllowedTextRotationValues.Count == 0)
                        {
                            config.AllowedTextRotationValues.Add(45);
                        }
                        config.RequireAllTexts ??= true;
                    }

                    if (task.SpecialCondition.ExcelMultiColumnSortConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelMultiColumnSortConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        if (config.HeaderRow <= 0)
                        {
                            config.HeaderRow = 1;
                        }
                        config.DataRange = string.IsNullOrWhiteSpace(config.DataRange) ? null : NormalizeExcelRangeAddress(config.DataRange);
                        config.KeyColumns = config.KeyColumns?
                            .Where(key => key != null)
                            .Select(key => new ExcelSortKeyConfig
                            {
                                HeaderName = string.IsNullOrWhiteSpace(key.HeaderName) ? null : NormalizePlainText(key.HeaderName),
                                Column = string.IsNullOrWhiteSpace(key.Column) ? null : key.Column.Trim().Replace("$", string.Empty).ToUpperInvariant(),
                                Descending = key.Descending ?? false
                            })
                            .Where(key => !string.IsNullOrWhiteSpace(key.HeaderName) || Regex.IsMatch(key.Column ?? string.Empty, "^[A-Z]{1,3}$", RegexOptions.CultureInvariant))
                            .ToList() ?? new List<ExcelSortKeyConfig>();
                    }

                    if (task.SpecialCondition.ExcelFreezePanesConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelFreezePanesConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? null : NormalizeSourceFile(config.SourceFile);
                        config.TopLeftCell = string.IsNullOrWhiteSpace(config.TopLeftCell) ? "A4" : NormalizeExcelCellAddress(config.TopLeftCell);
                        if (config.YSplit < 0)
                        {
                            config.YSplit = 3m;
                        }
                        if (config.XSplit < 0)
                        {
                            config.XSplit = 0m;
                        }
                        config.RequireNoColumnFreeze ??= true;
                    }

                    if (task.SpecialCondition.ExcelDocumentPropertyConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelDocumentPropertyConfig;
                        config.PropertyName = string.IsNullOrWhiteSpace(config.PropertyName) ? null : config.PropertyName.Trim();
                        config.ExpectedValue = string.IsNullOrWhiteSpace(config.ExpectedValue) ? null : NormalizePlainText(config.ExpectedValue);
                        config.SourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? "docProps/custom.xml" : NormalizeSourceFile(config.SourceFile);
                    }

                    if (task.SpecialCondition.ExcelPrintAreaConfig != null)
                    {
                        var config = task.SpecialCondition.ExcelPrintAreaConfig;
                        config.WorksheetName = string.IsNullOrWhiteSpace(config.WorksheetName) ? null : config.WorksheetName.Trim();
                        config.ExpectedRange = string.IsNullOrWhiteSpace(config.ExpectedRange) ? null : NormalizeExcelFormulaReference(config.ExpectedRange);
                        config.RequireExactRange ??= true;
                    }
                }
            }
        }

        private static void NormalizeConditionForPersistence(XmlGradingCondition condition)
        {
            condition.ConditionId = condition.ConditionId?.Trim() ?? string.Empty;
            condition.SourceFile = NormalizeSourceFile(condition.SourceFile);
            condition.ExpectedVariants = condition.ExpectedVariants?
                .Where(variant => variant != null)
                .Select(variant => new XmlExpectedVariant
                {
                    ExpectedValues = variant.ExpectedValues?
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Select(value => value.Trim())
                        .ToList() ?? new List<string>()
                })
                .Where(variant => variant.ExpectedValues.Count > 0)
                .ToList() ?? new List<XmlExpectedVariant>();
            condition.IgnoreAttributes = condition.IgnoreAttributes?
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();
            condition.CompareMode = string.IsNullOrWhiteSpace(condition.CompareMode)
                ? XmlGradingCompareModes.XmlContainsNormalized
                : condition.CompareMode.Trim();
            condition.MatchPolicy = string.IsNullOrWhiteSpace(condition.MatchPolicy)
                ? XmlGradingMatchPolicies.All
                : condition.MatchPolicy.Trim();
            if (condition.MinOccurrences <= 0)
            {
                condition.MinOccurrences = null;
            }
            if (condition.MaxOccurrences <= 0)
            {
                condition.MaxOccurrences = null;
            }
            condition.Feedback ??= new ConditionFeedback();
        }

        private static void ValidateRuleSetShell(GradingRuleSet ruleSet)
        {
            if (string.IsNullOrWhiteSpace(ruleSet.Subject))
            {
                throw new InvalidOperationException("Subject lÃ  báº¯t buá»™c.");
            }

            if (string.IsNullOrWhiteSpace(ruleSet.Version))
            {
                throw new InvalidOperationException("Version lÃ  báº¯t buá»™c.");
            }
        }

        private static void ValidateProjectShell(ProjectXmlRule project)
        {
            if (string.IsNullOrWhiteSpace(project.ProjectCode))
            {
                throw new InvalidOperationException("ProjectCode lÃ  báº¯t buá»™c.");
            }

            if (project.MaxScore <= 0)
            {
                throw new InvalidOperationException("Project maxScore pháº£i lá»›n hÆ¡n 0.");
            }
        }

        private static void ValidateTaskShell(TaskXmlRule task)
        {
            if (string.IsNullOrWhiteSpace(task.TaskId))
            {
                throw new InvalidOperationException("TaskId lÃ  báº¯t buá»™c.");
            }

            if (task.MaxScore <= 0)
            {
                throw new InvalidOperationException("Task maxScore pháº£i lá»›n hÆ¡n 0.");
            }

            if (task.SpecialCondition != null)
            {
                if (!SpecialConditionTypes.Supported.Contains(task.SpecialCondition.Type))
                {
                    throw new InvalidOperationException($"specialCondition.type khÃ´ng Ä‘Æ°á»£c há»— trá»£: {task.SpecialCondition.Type}.");
                }

                if (task.SpecialCondition.Score <= 0)
                {
                    throw new InvalidOperationException("specialCondition.score pháº£i lá»›n hÆ¡n 0.");
                }
            }
        }

        private static void ValidateConditionShell(XmlGradingCondition condition)
        {
            if (string.IsNullOrWhiteSpace(condition.ConditionId))
            {
                throw new InvalidOperationException("ConditionId lÃ  báº¯t buá»™c.");
            }

            if (condition.Score <= 0)
            {
                throw new InvalidOperationException("Condition score pháº£i lá»›n hÆ¡n 0.");
            }

            if (!IsSafeSourceFile(condition.SourceFile))
            {
                throw new InvalidOperationException("sourceFile pháº£i lÃ  Ä‘Æ°á»ng dáº«n XML an toÃ n trong Office package.");
            }

            if (condition.ExpectedVariants == null || condition.ExpectedVariants.Count == 0)
            {
                throw new InvalidOperationException("expectedVariants lÃ  báº¯t buá»™c.");
            }

            if (condition.ExpectedVariants.Any(variant => variant == null || variant.ExpectedValues.Count == 0))
            {
                throw new InvalidOperationException("Má»—i expectedVariant pháº£i cÃ³ Ã­t nháº¥t 1 expectedValues.");
            }

            if (!XmlGradingCompareModes.Supported.Contains(condition.CompareMode))
            {
                throw new InvalidOperationException($"compareMode khÃ´ng há»— trá»£: {condition.CompareMode}.");
            }

            if (!XmlGradingMatchPolicies.Supported.Contains(condition.MatchPolicy))
            {
                throw new InvalidOperationException($"matchPolicy khÃ´ng há»— trá»£: {condition.MatchPolicy}.");
            }
        }

        private static void EnsureMinOccurrencesIsValid(XmlGradingCondition condition)
        {
            if (string.Equals(condition.CompareMode, XmlGradingCompareModes.XmlMinOccurrences, StringComparison.OrdinalIgnoreCase)
                && (!condition.MinOccurrences.HasValue || condition.MinOccurrences.Value <= 0))
            {
                throw new InvalidOperationException("minOccurrences phai lon hon 0 khi compareMode la xmlMinOccurrences.");
            }

            if (condition.MaxOccurrences.HasValue
                && condition.MinOccurrences.HasValue
                && condition.MaxOccurrences.Value < condition.MinOccurrences.Value)
            {
                throw new InvalidOperationException("maxOccurrences phai lon hon hoac bang minOccurrences.");
            }
        }

        private static void EnsureSpecialConditionSupportedForSubject(
            SpecialCondition? specialCondition,
            string subject)
        {
            if (specialCondition == null || string.IsNullOrWhiteSpace(specialCondition.Type))
            {
                return;
            }

            if (!IsSpecialConditionSupportedForSubject(specialCondition.Type, subject))
            {
                throw new InvalidOperationException($"specialCondition.type {specialCondition.Type} khong ho tro cho subject {NormalizeKey(subject)}.");
            }
        }

        private static bool IsSpecialConditionSupportedForSubject(string specialConditionType, string subject)
        {
            var normalizedSubject = NormalizeKey(subject);

            return normalizedSubject switch
            {
                "word" => string.Equals(specialConditionType, SpecialConditionTypes.PictureBullet, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.InsertedImage, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ConvertTableToText, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.Hyperlink, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.SectionBreakBeforeText, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.PictureStyle, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.TextBoxContainsText, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.PageMargins, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.DocumentStyleSet, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.PageBorder, StringComparison.OrdinalIgnoreCase),
                "excel" => string.Equals(specialConditionType, SpecialConditionTypes.ExcelTableName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelWorksheetPageSetup, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelClearCellFormatting, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelDataModelImport, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelCompatibilityReport, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelMergedRange, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelCellHyperlink, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelIconSetConditionalFormatting, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelChartDataRange, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelChartStyle, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelTextReplacement, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelPrintTitles, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelNumberFormat, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelChartLegend, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelDefinedName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelFormulaReferences, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelNoConditionalFormatting, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelTextRotation, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelMultiColumnSort, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelFreezePanes, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelDocumentProperty, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(specialConditionType, SpecialConditionTypes.ExcelPrintArea, StringComparison.OrdinalIgnoreCase),
                "ppt" => false,
                "powerpoint" => false,
                _ => false
            };
        }

        private sealed class RequiredOfficeParts
        {
            public HashSet<string> XmlParts { get; } = new(StringComparer.OrdinalIgnoreCase);
            public HashSet<string> XmlPartPrefixes { get; } = new(StringComparer.OrdinalIgnoreCase);
            public bool ReadRelatedImages { get; set; }
            public bool ReadAllXmlParts { get; set; }
        }

        private sealed class OfficePackage
        {
            public HashSet<string> PartNames { get; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, string> XmlParts { get; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, byte[]> BinaryParts { get; } =
                new(StringComparer.OrdinalIgnoreCase);

            private readonly Dictionary<string, XDocument> _xmlDocuments =
                new(StringComparer.OrdinalIgnoreCase);

            private readonly Dictionary<string, Dictionary<string, string>> _relationshipMaps =
                new(StringComparer.OrdinalIgnoreCase);

            private readonly Dictionary<string, string> _sha256Hashes =
                new(StringComparer.OrdinalIgnoreCase);

            private readonly Dictionary<string, string> _perceptualHashes =
                new(StringComparer.OrdinalIgnoreCase);

            private string? _normalizedPartNameHaystack;
            private string? _normalizedXmlHaystack;

            public string GetNormalizedPartNameHaystack()
            {
                _normalizedPartNameHaystack ??= NormalizePlainText(string.Join(" ", PartNames));
                return _normalizedPartNameHaystack;
            }

            public string GetNormalizedXmlHaystack()
            {
                _normalizedXmlHaystack ??= NormalizePlainText(string.Join(" ", XmlParts.Values));
                return _normalizedXmlHaystack;
            }

            public bool TryGetXmlDocument(string sourceFile, out XDocument document, out string? errorMessage)
            {
                document = null!;
                errorMessage = null;
                var normalizedSourceFile = NormalizeSourceFile(sourceFile);

                if (_xmlDocuments.TryGetValue(normalizedSourceFile, out document!))
                {
                    return true;
                }

                if (!XmlParts.TryGetValue(normalizedSourceFile, out var xml))
                {
                    errorMessage = $"Khong tim thay {normalizedSourceFile} trong file hoc sinh.";
                    return false;
                }

                try
                {
                    document = XDocument.Parse(xml);
                    _xmlDocuments[normalizedSourceFile] = document;
                    return true;
                }
                catch (XmlException ex)
                {
                    errorMessage = $"Khong the phan tich XML {normalizedSourceFile}: {ex.Message}";
                    return false;
                }
            }

            public bool TryGetRelationships(string relsFile, out Dictionary<string, string> relationships, out string? errorMessage)
            {
                relationships = null!;
                errorMessage = null;
                var normalizedRelsFile = NormalizeSourceFile(relsFile);

                if (_relationshipMaps.TryGetValue(normalizedRelsFile, out relationships!))
                {
                    return true;
                }

                if (!TryGetXmlDocument(normalizedRelsFile, out var relsDocument, out errorMessage))
                {
                    return false;
                }

                XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
                relationships = relsDocument
                    .Descendants(rel + "Relationship")
                    .Where(relationship => !string.IsNullOrWhiteSpace(relationship.Attribute("Id")?.Value))
                    .GroupBy(
                        relationship => relationship.Attribute("Id")!.Value,
                        StringComparer.Ordinal)
                    .ToDictionary(
                        group => group.Key,
                        group => group.First().Attribute("Target")?.Value ?? string.Empty,
                        StringComparer.Ordinal);

                _relationshipMaps[normalizedRelsFile] = relationships;
                return true;
            }

            public string GetSha256(string imagePath, byte[] imageBytes)
            {
                var normalizedPath = NormalizeSourceFile(imagePath);
                if (_sha256Hashes.TryGetValue(normalizedPath, out var hash))
                {
                    return hash;
                }

                hash = ImageHashUtility.ComputeSha256(imageBytes);
                _sha256Hashes[normalizedPath] = hash;
                return hash;
            }

            public string GetPerceptualHash(string imagePath, byte[] imageBytes)
            {
                var normalizedPath = NormalizeSourceFile(imagePath);
                if (_perceptualHashes.TryGetValue(normalizedPath, out var hash))
                {
                    return hash;
                }

                hash = ImageHashUtility.ComputePerceptualHash(imageBytes);
                _perceptualHashes[normalizedPath] = hash;
                return hash;
            }
        }

        private static string BuildIgnoreAttributesKey(IReadOnlyList<string>? ignoreAttributes)
        {
            if (ignoreAttributes == null || ignoreAttributes.Count == 0)
            {
                return string.Empty;
            }

            return string.Join(
                "|",
                ignoreAttributes
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(NormalizeIgnoreAttributeToken)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
        }

        private sealed class XmlEvaluationCache
        {
            private readonly Dictionary<string, string> _normalizedActualXml = new(StringComparer.Ordinal);
            private readonly Dictionary<string, string> _normalizedExpectedXml = new(StringComparer.Ordinal);

            public string GetNormalizedActual(
                string sourceFile,
                string value,
                IReadOnlyList<string>? ignoreAttributes)
            {
                var key = $"{sourceFile}::{BuildIgnoreAttributesKey(ignoreAttributes)}";
                if (_normalizedActualXml.TryGetValue(key, out var normalized))
                {
                    return normalized;
                }

                normalized = NormalizeXmlForComparison(value, ignoreAttributes);
                _normalizedActualXml[key] = normalized;
                return normalized;
            }

            public string GetNormalizedExpected(
                string value,
                IReadOnlyList<string>? ignoreAttributes)
            {
                var key = $"{BuildIgnoreAttributesKey(ignoreAttributes)}::{value}";
                if (_normalizedExpectedXml.TryGetValue(key, out var normalized))
                {
                    return normalized;
                }

                normalized = NormalizeXmlForComparison(value, ignoreAttributes);
                _normalizedExpectedXml[key] = normalized;
                return normalized;
            }
        }

        private static OfficePackage ReadOfficePackage(Stream studentFile, RequiredOfficeParts? requiredParts = null)
        {
            if (studentFile.CanSeek)
            {
                studentFile.Position = 0;
            }

            using var archive = new ZipArchive(
                studentFile,
                ZipArchiveMode.Read,
                leaveOpen: true);

            var package = new OfficePackage();
            var entriesByPath = archive.Entries
                .Where(entry => !string.IsNullOrWhiteSpace(entry.FullName))
                .GroupBy(entry => NormalizeSourceFile(entry.FullName), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            foreach (var path in entriesByPath.Keys)
            {
                package.PartNames.Add(path);
            }

            if (requiredParts == null || requiredParts.XmlParts.Count == 0 || requiredParts.ReadAllXmlParts)
            {
                foreach (var entry in archive.Entries)
                {
                    var normalizedPath = NormalizeSourceFile(entry.FullName);

                    if (string.IsNullOrWhiteSpace(normalizedPath))
                    {
                        continue;
                    }

                    if (IsXmlLikePart(normalizedPath))
                    {
                        package.XmlParts[normalizedPath] = ReadEntryText(entry);
                    }
                    else if (IsSupportedImage(normalizedPath))
                    {
                        package.BinaryParts[normalizedPath] = ReadEntryBytes(entry);
                    }
                }

                return package;
            }

            foreach (var xmlPart in requiredParts.XmlParts)
            {
                var normalizedPath = NormalizeSourceFile(xmlPart);
                if (entriesByPath.TryGetValue(normalizedPath, out var entry))
                {
                    package.XmlParts[normalizedPath] = ReadEntryText(entry);
                }
            }

            if (requiredParts.XmlPartPrefixes.Count > 0)
            {
                foreach (var (path, entry) in entriesByPath)
                {
                    if (!IsXmlLikePart(path) || package.XmlParts.ContainsKey(path))
                    {
                        continue;
                    }

                    if (requiredParts.XmlPartPrefixes.Any(prefix =>
                        path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                    {
                        package.XmlParts[path] = ReadEntryText(entry);
                    }
                }
            }

            if (requiredParts.ReadRelatedImages)
            {
                foreach (var imagePath in CollectRelatedImageParts(package))
                {
                    if (entriesByPath.TryGetValue(imagePath, out var entry))
                    {
                        package.BinaryParts[imagePath] = ReadEntryBytes(entry);
                    }
                }
            }

            return package;
        }

        private static bool IsXmlLikePart(string path)
        {
            return path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
        }

        private static string ReadEntryText(ZipArchiveEntry entry)
        {
            using var entryStream = entry.Open();
            using var reader = new StreamReader(
                entryStream,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true);

            return reader.ReadToEnd();
        }

        private static byte[] ReadEntryBytes(ZipArchiveEntry entry)
        {
            using var entryStream = entry.Open();
            using var memoryStream = new MemoryStream();
            entryStream.CopyTo(memoryStream);
            return memoryStream.ToArray();
        }

        private static RequiredOfficeParts CollectRequiredOfficeParts(ProjectXmlRule projectRule)
        {
            var requiredParts = new RequiredOfficeParts();

            void AddXmlPart(string? path, string fallback)
            {
                var normalizedPath = string.IsNullOrWhiteSpace(path)
                    ? fallback
                    : NormalizeSourceFile(path);

                if (!string.IsNullOrWhiteSpace(normalizedPath))
                {
                    requiredParts.XmlParts.Add(normalizedPath);
                }
            }

            void AddXmlPartIfProvided(string? path)
            {
                var normalizedPath = NormalizeSourceFile(path ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(normalizedPath))
                {
                    requiredParts.XmlParts.Add(normalizedPath);
                }
            }

            void AddXmlPrefix(string prefix)
            {
                var normalizedPrefix = NormalizeSourceFile(prefix);
                if (!string.IsNullOrWhiteSpace(normalizedPrefix))
                {
                    requiredParts.XmlPartPrefixes.Add(normalizedPrefix.TrimEnd('/') + "/");
                }
            }

            void AddExcelWorkbookParts()
            {
                AddXmlPart("xl/workbook.xml", "xl/workbook.xml");
                AddXmlPart("xl/_rels/workbook.xml.rels", "xl/_rels/workbook.xml.rels");
            }

            void AddExcelWorksheetParts(string? sourceFile)
            {
                AddExcelWorkbookParts();
                if (!string.IsNullOrWhiteSpace(sourceFile))
                {
                    AddXmlPartIfProvided(sourceFile);
                    AddXmlPartIfProvided(GetRelationshipPartPath(sourceFile));
                    return;
                }

                AddXmlPrefix("xl/worksheets");
            }

            foreach (var task in projectRule.Tasks)
            {
                foreach (var condition in task.Conditions)
                {
                    AddXmlPart(condition.SourceFile, "word/document.xml");
                }

                var specialCondition = task.SpecialCondition;
                if (specialCondition == null || string.IsNullOrWhiteSpace(specialCondition.Type))
                {
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureBullet, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart("word/document.xml", "word/document.xml");
                    AddXmlPart("word/numbering.xml", "word/numbering.xml");
                    AddXmlPart("word/_rels/numbering.xml.rels", "word/_rels/numbering.xml.rels");
                    requiredParts.ReadRelatedImages = true;
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.InsertedImage, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.ImageInsertConfig?.SourceFile, "word/document.xml");
                    AddXmlPart(specialCondition.ImageInsertConfig?.RelsFile, "word/_rels/document.xml.rels");
                    requiredParts.ReadRelatedImages = true;
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ConvertTableToText, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.ConvertTableToTextConfig?.SourceFile, "word/document.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.Hyperlink, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.HyperlinkConfig?.SourceFile, "word/document.xml");
                    AddXmlPart(specialCondition.HyperlinkConfig?.RelsFile, "word/_rels/document.xml.rels");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.SectionBreakBeforeText, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.SectionBreakBeforeTextConfig?.SourceFile, "word/document.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureStyle, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.PictureStyleConfig?.SourceFile, "word/document.xml");
                    if (!string.IsNullOrWhiteSpace(specialCondition.PictureStyleConfig?.ImageHash))
                    {
                        AddXmlPart(specialCondition.PictureStyleConfig?.RelsFile, "word/_rels/document.xml.rels");
                        requiredParts.ReadRelatedImages = true;
                    }
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.TextBoxContainsText, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.TextBoxContainsTextConfig?.SourceFile, "word/document.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.PageMargins, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.PageMarginsConfig?.SourceFile, "word/document.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.DocumentStyleSet, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.DocumentStyleSetConfig?.SourceFile, "word/styles.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.PageBorder, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.PageBorderConfig?.SourceFile, "word/document.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTableName, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelTableNameConfig?.SourceFile);
                    AddXmlPrefix("xl/tables");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelWorksheetPageSetup, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelWorksheetPageSetupConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelClearCellFormatting, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelClearCellFormattingConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDataModelImport, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorkbookParts();
                    AddXmlPart("xl/connections.xml", "xl/connections.xml");
                    AddXmlPrefix("xl/queryTables");
                    AddXmlPrefix("xl/tables");
                    AddXmlPrefix("xl/model");
                    AddXmlPrefix("customXml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelCompatibilityReport, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(null);
                    AddXmlPart("xl/sharedStrings.xml", "xl/sharedStrings.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelMergedRange, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelMergedRangeConfig?.SourceFile);
                    if (specialCondition.ExcelMergedRangeConfig?.RequireNoHorizontalCenter == true)
                    {
                        AddXmlPart("xl/styles.xml", "xl/styles.xml");
                    }
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelCellHyperlink, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelCellHyperlinkConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelIconSetConditionalFormatting, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelIconSetConditionalFormattingConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartDataRange, StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(specialCondition.ExcelChartDataRangeConfig?.ChartSourceFile))
                    {
                        AddXmlPartIfProvided(specialCondition.ExcelChartDataRangeConfig.ChartSourceFile);
                    }
                    else
                    {
                        AddXmlPrefix("xl/charts");
                    }
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartStyle, StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(specialCondition.ExcelChartStyleConfig?.ChartSourceFile))
                    {
                        AddXmlPartIfProvided(specialCondition.ExcelChartStyleConfig.ChartSourceFile);
                    }
                    else
                    {
                        AddXmlPrefix("xl/charts");
                    }

                    if (!string.IsNullOrWhiteSpace(specialCondition.ExcelChartStyleConfig?.StyleSourceFile))
                    {
                        AddXmlPartIfProvided(specialCondition.ExcelChartStyleConfig.StyleSourceFile);
                    }
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTextReplacement, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorkbookParts();
                    AddXmlPart("xl/sharedStrings.xml", "xl/sharedStrings.xml");
                    AddExcelWorksheetParts(specialCondition.ExcelTextReplacementConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelPrintTitles, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorkbookParts();
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelNumberFormat, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelNumberFormatConfig?.SourceFile);
                    AddXmlPart("xl/styles.xml", "xl/styles.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartLegend, StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(specialCondition.ExcelChartLegendConfig?.ChartSourceFile))
                    {
                        AddXmlPartIfProvided(specialCondition.ExcelChartLegendConfig.ChartSourceFile);
                    }
                    else
                    {
                        AddXmlPrefix("xl/charts");
                    }
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDefinedName, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorkbookParts();
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelFormulaReferences, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelFormulaReferencesConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelNoConditionalFormatting, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelNoConditionalFormattingConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTextRotation, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelTextRotationConfig?.SourceFile);
                    AddXmlPart("xl/sharedStrings.xml", "xl/sharedStrings.xml");
                    AddXmlPart("xl/styles.xml", "xl/styles.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelMultiColumnSort, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelMultiColumnSortConfig?.SourceFile);
                    AddXmlPart("xl/sharedStrings.xml", "xl/sharedStrings.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelFreezePanes, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorksheetParts(specialCondition.ExcelFreezePanesConfig?.SourceFile);
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDocumentProperty, StringComparison.OrdinalIgnoreCase))
                {
                    AddXmlPart(specialCondition.ExcelDocumentPropertyConfig?.SourceFile, "docProps/custom.xml");
                    continue;
                }

                if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelPrintArea, StringComparison.OrdinalIgnoreCase))
                {
                    AddExcelWorkbookParts();
                    continue;
                }
            }

            return requiredParts;
        }

        private static IEnumerable<string> CollectRelatedImageParts(OfficePackage package)
        {
            var imageParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var (relsPath, relsXml) in package.XmlParts
                .Where(part => part.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
            {
                XDocument relationshipsDocument;
                try
                {
                    relationshipsDocument = XDocument.Parse(relsXml);
                }
                catch (XmlException)
                {
                    continue;
                }

                var sourcePart = ResolveRelationshipsSourcePart(relsPath);
                XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

                foreach (var relationship in relationshipsDocument.Descendants(rel + "Relationship"))
                {
                    var target = relationship.Attribute("Target")?.Value;
                    if (string.IsNullOrWhiteSpace(target))
                    {
                        continue;
                    }

                    var imagePath = ResolveRelationshipTarget(sourcePart, target);
                    if (IsSupportedImage(imagePath))
                    {
                        imageParts.Add(imagePath);
                    }
                }
            }

            return imageParts;
        }

        private static string ResolveRelationshipsSourcePart(string relsPath)
        {
            var normalizedPath = NormalizeSourceFile(relsPath);
            const string relsMarker = "/_rels/";

            var markerIndex = normalizedPath.LastIndexOf(relsMarker, StringComparison.OrdinalIgnoreCase);
            if (markerIndex < 0 || !normalizedPath.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                return normalizedPath;
            }

            var directory = normalizedPath[..markerIndex];
            var fileName = normalizedPath[(markerIndex + relsMarker.Length)..^5];

            return string.IsNullOrWhiteSpace(directory)
                ? fileName
                : $"{directory}/{fileName}";
        }

        private static bool IsSupportedImage(string path)
        {
            return path.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".gif", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".wmf", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".emf", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".webp", StringComparison.OrdinalIgnoreCase);
        }

        private static XmlConditionEvaluationResult EvaluateCondition(
            XmlGradingCondition condition,
            OfficePackage package,
            XmlEvaluationCache cache)
        {
            var compareMode = string.IsNullOrWhiteSpace(condition.CompareMode)
                ? XmlGradingCompareModes.XmlContainsNormalized
                : condition.CompareMode.Trim();

            var matchPolicy = string.IsNullOrWhiteSpace(condition.MatchPolicy)
                ? XmlGradingMatchPolicies.All
                : condition.MatchPolicy.Trim();

            var result = new XmlConditionEvaluationResult
            {
                ConditionId = condition.ConditionId,
                SourceFile = NormalizeSourceFile(condition.SourceFile),
                CompareMode = compareMode,
                MatchPolicy = matchPolicy,
                MaxConditionScore = condition.Score,
                Feedback = condition.Feedback ?? new ConditionFeedback()
            };

            if (!package.XmlParts.TryGetValue(result.SourceFile, out var actualXml))
            {
                result.MissingExpectedValues.AddRange(
                    (condition.ExpectedVariants ?? new List<XmlExpectedVariant>())
                        .Where(variant => variant != null)
                        .SelectMany(variant => variant.ExpectedValues)
                        .Distinct(StringComparer.Ordinal));
                if (string.IsNullOrWhiteSpace(result.Feedback.ErrorMessage))
                {
                    result.Feedback.ErrorMessage = $"KhÃ´ng tÃ¬m tháº¥y XML part {result.SourceFile} trong file há»c sinh.";
                }

                return result;
            }

            List<ExpectedMatchResult>? bestMatches = null;
            foreach (var variant in (condition.ExpectedVariants ?? new List<XmlExpectedVariant>()).Where(variant => variant != null))
            {
                var matches = EvaluateExpectedValues(result.SourceFile, actualXml, variant.ExpectedValues, compareMode, matchPolicy, condition.IgnoreAttributes, condition.MinOccurrences, condition.MaxOccurrences, cache);
                bestMatches ??= matches;

                if (ApplyMatchPolicy(matches, matchPolicy))
                {
                    result.IsPassed = true;
                    result.ScoreAwarded = condition.Score;
                    result.MatchedExpectedValues = matches.Where(match => match.IsMatched).Select(match => match.ExpectedValue).ToList();
                    result.MissingExpectedValues = matches.Where(match => !match.IsMatched).Select(match => match.ExpectedValue).ToList();
                    return result;
                }

                if (matches.Count(match => match.IsMatched) > bestMatches.Count(match => match.IsMatched))
                {
                    bestMatches = matches;
                }
            }

            result.IsPassed = false;
            result.ScoreAwarded = 0m;
            result.MatchedExpectedValues = bestMatches?.Where(match => match.IsMatched).Select(match => match.ExpectedValue).ToList() ?? new List<string>();
            result.MissingExpectedValues = bestMatches?.Where(match => !match.IsMatched).Select(match => match.ExpectedValue).ToList() ?? new List<string>();

            return result;
        }

        private static List<ExpectedMatchResult> EvaluateExpectedValues(
            string sourceFile,
            string actualXml,
            IReadOnlyList<string> expectedValues,
            string compareMode,
            string matchPolicy,
            IReadOnlyList<string> ignoreAttributes,
            int? minOccurrences,
            int? maxOccurrences,
            XmlEvaluationCache cache)
        {
            return string.Equals(matchPolicy, XmlGradingMatchPolicies.Ordered, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(compareMode, XmlGradingCompareModes.XmlEquivalentWholeFile, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(compareMode, XmlGradingCompareModes.XmlMinOccurrences, StringComparison.OrdinalIgnoreCase)
                ? MatchExpectedOrdered(sourceFile, actualXml, expectedValues, compareMode, ignoreAttributes, cache)
                : expectedValues.Select(expected => MatchExpected(sourceFile, actualXml, expected, compareMode, ignoreAttributes, minOccurrences, maxOccurrences, cache)).ToList();
        }

        /// <summary>
        /// Káº¿t quáº£ ná»™i bá»™ khi Ä‘Ã¡nh giÃ¡ 1 Special Condition (khÃ´ng phÆ¡i ra ngoÃ i API,
        /// chá»‰ dÃ¹ng Ä‘á»ƒ build message cho taskResult.Details/Errors).
        /// </summary>
        private sealed class SpecialConditionEvalOutcome
        {
            public bool IsPassed { get; set; }
            public string Message { get; set; } = string.Empty;
        }

        private static SpecialConditionEvalOutcome EvaluateTaskSpecialCondition(SpecialCondition specialCondition, OfficePackage package, string subject)
        {
            if (!IsSpecialConditionSupportedForSubject(specialCondition.Type, subject))
            {
                return new SpecialConditionEvalOutcome
                {
                    IsPassed = false,
                    Message = $"Special condition {specialCondition.Type} khong ho tro cho subject {NormalizeKey(subject)}."
                };
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureBullet, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluatePictureBullet(specialCondition.Config, package);
            }

            // Má»šI
            if (string.Equals(specialCondition.Type, SpecialConditionTypes.InsertedImage, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateInsertedImage(specialCondition.ImageInsertConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ConvertTableToText, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateConvertTableToText(specialCondition.ConvertTableToTextConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.Hyperlink, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateHyperlink(specialCondition.HyperlinkConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.SectionBreakBeforeText, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateSectionBreakBeforeText(specialCondition.SectionBreakBeforeTextConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureStyle, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluatePictureStyle(specialCondition.PictureStyleConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.TextBoxContainsText, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateTextBoxContainsText(specialCondition.TextBoxContainsTextConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PageMargins, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluatePageMargins(specialCondition.PageMarginsConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.DocumentStyleSet, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateDocumentStyleSet(specialCondition.DocumentStyleSetConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PageBorder, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluatePageBorder(specialCondition.PageBorderConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTableName, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelTableName(specialCondition.ExcelTableNameConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelWorksheetPageSetup, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelWorksheetPageSetup(specialCondition.ExcelWorksheetPageSetupConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelClearCellFormatting, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelClearCellFormatting(specialCondition.ExcelClearCellFormattingConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDataModelImport, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelDataModelImport(specialCondition.ExcelDataModelImportConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelCompatibilityReport, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelCompatibilityReport(specialCondition.ExcelCompatibilityReportConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelMergedRange, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelMergedRange(specialCondition.ExcelMergedRangeConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelCellHyperlink, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelCellHyperlink(specialCondition.ExcelCellHyperlinkConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelIconSetConditionalFormatting, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelIconSetConditionalFormatting(specialCondition.ExcelIconSetConditionalFormattingConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartDataRange, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelChartDataRange(specialCondition.ExcelChartDataRangeConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartStyle, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelChartStyle(specialCondition.ExcelChartStyleConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTextReplacement, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelTextReplacement(specialCondition.ExcelTextReplacementConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelPrintTitles, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelPrintTitles(specialCondition.ExcelPrintTitlesConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelNumberFormat, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelNumberFormat(specialCondition.ExcelNumberFormatConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartLegend, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelChartLegend(specialCondition.ExcelChartLegendConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDefinedName, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelDefinedName(specialCondition.ExcelDefinedNameConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelFormulaReferences, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelFormulaReferences(specialCondition.ExcelFormulaReferencesConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelNoConditionalFormatting, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelNoConditionalFormatting(specialCondition.ExcelNoConditionalFormattingConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTextRotation, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelTextRotation(specialCondition.ExcelTextRotationConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelMultiColumnSort, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelMultiColumnSort(specialCondition.ExcelMultiColumnSortConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelFreezePanes, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelFreezePanes(specialCondition.ExcelFreezePanesConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDocumentProperty, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelDocumentProperty(specialCondition.ExcelDocumentPropertyConfig, package);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelPrintArea, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateExcelPrintArea(specialCondition.ExcelPrintAreaConfig, package);
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = false,
                Message = $"Special condition khÃ´ng Ä‘Æ°á»£c há»— trá»£: {specialCondition.Type}."
            };
        }

        /// <summary>
        /// So sÃ¡nh áº£nh thá»±c táº¿ vá»›i áº£nh chuáº©n (expectedHash/expectedPerceptualHash).
        /// Æ¯u tiÃªn perceptual hash (chá»‹u Ä‘Æ°á»£c Word nÃ©n láº¡i JPEG khi save); náº¿u
        /// ruleset chÆ°a cÃ³ perceptualHash (dá»¯ liá»‡u cÅ©), hoáº·c áº£nh khÃ´ng decode
        /// Ä‘Æ°á»£c (vd .wmf/.emf), fallback vá» so SHA-256 tuyá»‡t Ä‘á»‘i nhÆ° trÆ°á»›c.
        /// </summary>
        private static bool IsImageMatch(byte[] imageBytes, string actualSha256Hash, string expectedSha256Hash, string? expectedPerceptualHash)
        {
            if (!string.IsNullOrWhiteSpace(expectedPerceptualHash))
            {
                try
                {
                    var actualPerceptualHash = ImageHashUtility.ComputePerceptualHash(imageBytes);
                    return ImageHashUtility.IsPerceptuallySimilar(actualPerceptualHash, expectedPerceptualHash, PerceptualHashThreshold);
                }
                catch
                {
                    // áº¢nh khÃ´ng decode Ä‘Æ°á»£c báº±ng ImageSharp (vd .wmf/.emf) -> fallback SHA-256
                    return string.Equals(actualSha256Hash, expectedSha256Hash, StringComparison.OrdinalIgnoreCase);
                }
            }

            return string.Equals(actualSha256Hash, expectedSha256Hash, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsImageMatch(
            OfficePackage package,
            string imagePath,
            byte[] imageBytes,
            string expectedSha256Hash,
            string? expectedPerceptualHash)
        {
            var actualSha256Hash = package.GetSha256(imagePath, imageBytes);

            if (!string.IsNullOrWhiteSpace(expectedPerceptualHash))
            {
                try
                {
                    var actualPerceptualHash = package.GetPerceptualHash(imagePath, imageBytes);
                    return ImageHashUtility.IsPerceptuallySimilar(actualPerceptualHash, expectedPerceptualHash, PerceptualHashThreshold);
                }
                catch
                {
                    return string.Equals(actualSha256Hash, expectedSha256Hash, StringComparison.OrdinalIgnoreCase);
                }
            }

            return string.Equals(actualSha256Hash, expectedSha256Hash, StringComparison.OrdinalIgnoreCase);
        }

        private sealed class ParagraphTextSnapshot
        {
            public string Text { get; init; } = string.Empty;
            public int TabCount { get; init; }
        }

        private sealed class InsertedImageCandidate
        {
            public XElement Drawing { get; init; } = null!;
            public int ParagraphIndex { get; init; }
        }

        private sealed class ExcelWorksheetInfo
        {
            public string Name { get; init; } = string.Empty;
            public string RelationshipId { get; init; } = string.Empty;
            public string SourceFile { get; init; } = string.Empty;
        }

        private static SpecialConditionEvalOutcome EvaluateExcelMergedRange(
            ExcelMergedRangeConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Merged Range (excelMergedRangeConfig trong).");
            }

            var expectedRange = NormalizeExcelRangeAddress(config.Range);
            if (string.IsNullOrWhiteSpace(expectedRange))
            {
                return Fail("excelMergedRangeConfig.range khong hop le. Vi du hop le: A1:E1.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return Fail(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {worksheetPath} trong file hoc sinh.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var hasMerge = document
                .Descendants(x + "mergeCell")
                .Any(mergeCell => string.Equals(
                    NormalizeExcelRangeAddress(mergeCell.Attribute("ref")?.Value),
                    expectedRange,
                    StringComparison.OrdinalIgnoreCase));

            if (!hasMerge)
            {
                return Fail($"Khong tim thay merged range {expectedRange} tren {worksheetPath}.");
            }

            if (config.RequireNoHorizontalCenter == true
                && TryGetHorizontallyCenteredCellInRange(package, document, x, expectedRange, out var centeredCell))
            {
                return Fail($"O {centeredCell} trong merged range {expectedRange} dang can giua ngang. Task nay can Merge Across/giu nguyen alignment, khong dung Merge & Center.");
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Worksheet {worksheetPath} co merged range {expectedRange}."
            };
        }

        private static bool TryGetHorizontallyCenteredCellInRange(
            OfficePackage package,
            XDocument worksheetDocument,
            XNamespace worksheetNamespace,
            string range,
            out string cellAddress)
        {
            cellAddress = string.Empty;

            if (!TryParseExcelRange(range, out var startColumn, out var startRow, out var endColumn, out var endRow))
            {
                return false;
            }

            if (!package.TryGetXmlDocument("xl/styles.xml", out var stylesDocument, out _))
            {
                return false;
            }

            var horizontallyCenteredStyleIds = GetHorizontallyCenteredStyleIds(stylesDocument);
            if (horizontallyCenteredStyleIds.Count == 0)
            {
                return false;
            }

            var cellsByAddress = worksheetDocument
                .Descendants(worksheetNamespace + "c")
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Attribute("r")?.Value))
                .ToDictionary(
                    cell => NormalizeExcelCellAddress(cell.Attribute("r")!.Value),
                    cell => cell,
                    StringComparer.OrdinalIgnoreCase);

            for (var row = startRow; row <= endRow; row++)
            {
                for (var column = startColumn; column <= endColumn; column++)
                {
                    var address = $"{GetExcelColumnName(column)}{row}";
                    if (!cellsByAddress.TryGetValue(address, out var cell))
                    {
                        continue;
                    }

                    var styleText = cell.Attribute("s")?.Value;
                    if (int.TryParse(styleText, out var styleId)
                        && horizontallyCenteredStyleIds.Contains(styleId))
                    {
                        cellAddress = address;
                        return true;
                    }
                }
            }

            return false;
        }

        private static HashSet<int> GetHorizontallyCenteredStyleIds(XDocument stylesDocument)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var result = new HashSet<int>();
            var cellXfs = stylesDocument
                .Root?
                .Element(x + "cellXfs")?
                .Elements(x + "xf")
                .ToList() ?? new List<XElement>();

            for (var index = 0; index < cellXfs.Count; index++)
            {
                var horizontal = cellXfs[index]
                    .Element(x + "alignment")?
                    .Attribute("horizontal")?
                    .Value;

                if (string.Equals(horizontal, "center", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(horizontal, "centerContinuous", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(index);
                }
            }

            return result;
        }

        private static SpecialConditionEvalOutcome EvaluateExcelCellHyperlink(
            ExcelCellHyperlinkConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Cell Hyperlink (excelCellHyperlinkConfig trong).");
            }

            var expectedCell = NormalizeExcelCellAddress(config.Cell);
            if (string.IsNullOrWhiteSpace(expectedCell))
            {
                return Fail("excelCellHyperlinkConfig.cell khong hop le. Vi du hop le: B13.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return Fail(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {worksheetPath} trong file hoc sinh.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var hyperlink = document
                .Descendants(x + "hyperlink")
                .FirstOrDefault(item => string.Equals(
                    NormalizeExcelRangeAddress(item.Attribute("ref")?.Value),
                    expectedCell,
                    StringComparison.OrdinalIgnoreCase));

            if (hyperlink == null)
            {
                return Fail($"Khong tim thay hyperlink tai o {expectedCell}.");
            }

            if (!string.IsNullOrWhiteSpace(config.Location)
                && !string.Equals(hyperlink.Attribute("location")?.Value?.Trim(), config.Location.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return Fail($"Hyperlink tai {expectedCell} co location '{hyperlink.Attribute("location")?.Value ?? "(rong)"}', can '{config.Location}'.");
            }

            if (!string.IsNullOrWhiteSpace(config.Display)
                && !string.Equals(hyperlink.Attribute("display")?.Value?.Trim(), config.Display.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return Fail($"Hyperlink tai {expectedCell} co display khong dung.");
            }

            if (!string.IsNullOrWhiteSpace(config.Target))
            {
                var relationshipId = hyperlink.Attribute(r + "id")?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId))
                {
                    return Fail($"Hyperlink tai {expectedCell} khong co relationship id de kiem tra target.");
                }

                var relsPath = GetRelationshipPartPath(worksheetPath);
                if (!package.TryGetRelationships(relsPath, out var relationships, out var relsError))
                {
                    return Fail(relsError ?? $"Khong doc duoc relationships {relsPath}.");
                }

                if (!relationships.TryGetValue(relationshipId, out var target)
                    || !string.Equals(target.Trim(), config.Target.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return Fail($"Hyperlink tai {expectedCell} khong tro toi target '{config.Target}'.");
                }
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"O {expectedCell} tren {worksheetPath} co hyperlink dung cau hinh."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelIconSetConditionalFormatting(
            ExcelIconSetConditionalFormattingConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Icon Set Conditional Formatting (excelIconSetConditionalFormattingConfig trong).");
            }

            var expectedRange = NormalizeExcelRangeAddress(config.Range);
            if (string.IsNullOrWhiteSpace(expectedRange))
            {
                return Fail("excelIconSetConditionalFormattingConfig.range khong hop le. Vi du hop le: C4:C11.");
            }

            var expectedIconSet = string.IsNullOrWhiteSpace(config.IconSet)
                ? "3Flags"
                : config.IconSet.Trim();

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return Fail(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {worksheetPath} trong file hoc sinh.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var matchingFormat = document
                .Descendants(x + "conditionalFormatting")
                .Where(item => RangeListContains(item.Attribute("sqref")?.Value, expectedRange))
                .SelectMany(item => item.Descendants(x + "iconSet"))
                .FirstOrDefault(iconSet => string.Equals(
                    iconSet.Attribute("iconSet")?.Value,
                    expectedIconSet,
                    StringComparison.OrdinalIgnoreCase));

            if (matchingFormat == null)
            {
                return Fail($"Khong tim thay icon set '{expectedIconSet}' tren range {expectedRange}.");
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Range {expectedRange} tren {worksheetPath} co icon set '{expectedIconSet}'."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelChartDataRange(
            ExcelChartDataRangeConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Chart Data Range (excelChartDataRangeConfig trong).");
            }

            var expectedCategoryRange = NormalizeExcelFormulaReference(config.ExpectedCategoryRange);
            var expectedValueRanges = config.ExpectedValueRanges?
                .Select(NormalizeExcelFormulaReference)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();
            var expectedValueRange = NormalizeExcelFormulaReference(config.ExpectedValueRange);
            if (!string.IsNullOrWhiteSpace(expectedValueRange)
                && !expectedValueRanges.Contains(expectedValueRange, StringComparer.OrdinalIgnoreCase))
            {
                expectedValueRanges.Insert(0, expectedValueRange);
            }
            var expectedSeriesNames = config.ExpectedSeriesNames?
                .Select(NormalizePlainText)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();
            var expectedCategoryText = NormalizePlainText(config.ExpectedCategoryText);

            var chartParts = ResolveExcelChartParts(package, config.ChartSourceFile);
            if (chartParts.Count == 0)
            {
                return Fail("Khong tim thay chart XML can kiem tra trong workbook.");
            }

            foreach (var chartPath in chartParts)
            {
                if (!package.TryGetXmlDocument(chartPath, out var chartDocument, out _))
                {
                    continue;
                }

                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                var formulas = chartDocument
                    .Descendants(c + "f")
                    .Select(item => NormalizeExcelFormulaReference(item.Value))
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .ToList();
                var series = chartDocument
                    .Descendants(c + "ser")
                    .Select(item => new
                    {
                        CategoryFormulas = item
                            .Elements(c + "cat")
                            .Descendants(c + "f")
                            .Select(formula => NormalizeExcelFormulaReference(formula.Value))
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .ToList(),
                        ValueFormulas = item
                            .Elements(c + "val")
                            .Descendants(c + "f")
                            .Select(formula => NormalizeExcelFormulaReference(formula.Value))
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .ToList(),
                        SeriesNames = item
                            .Elements(c + "tx")
                            .Descendants(c + "v")
                            .Select(value => NormalizePlainText(value.Value))
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .ToList()
                    })
                    .ToList();

                var categoryOk = string.IsNullOrWhiteSpace(expectedCategoryRange)
                    || formulas.Any(value => string.Equals(value, expectedCategoryRange, StringComparison.OrdinalIgnoreCase));
                var valueOk = expectedValueRanges.Count == 0
                    || expectedValueRanges.All(expected =>
                        formulas.Any(value => string.Equals(value, expected, StringComparison.OrdinalIgnoreCase)));
                var matchingSeries = series
                    .Where(item => item.ValueFormulas.Count > 0 || item.CategoryFormulas.Count > 0)
                    .ToList();
                var matchingSeriesOk = matchingSeries.Count == 0
                    || expectedValueRanges.Count == 0
                    || expectedValueRanges.All(expected =>
                        matchingSeries.Any(item =>
                            item.ValueFormulas.Any(value => string.Equals(value, expected, StringComparison.OrdinalIgnoreCase))
                            && (string.IsNullOrWhiteSpace(expectedCategoryRange)
                                || item.CategoryFormulas.Any(value => string.Equals(value, expectedCategoryRange, StringComparison.OrdinalIgnoreCase)))));
                var noExtraSeriesOk = config.RequireNoExtraSeries != true
                    || expectedValueRanges.Count == 0
                    || matchingSeries.Count == 0
                    || (matchingSeries.Count == expectedValueRanges.Count
                        && matchingSeries.All(item => item.ValueFormulas.Any(value =>
                            expectedValueRanges.Contains(value, StringComparer.OrdinalIgnoreCase))));
                var seriesNamesOk = expectedSeriesNames.Count == 0
                    || expectedSeriesNames.All(expected =>
                        matchingSeries.Any(item => item.SeriesNames.Any(value =>
                            string.Equals(value, expected, StringComparison.OrdinalIgnoreCase))));
                var pointCountOk = !config.ExpectedPointCount.HasValue
                    || chartDocument.Descendants(c + "ptCount").Any(item =>
                        int.TryParse(item.Attribute("val")?.Value, out var count) && count == config.ExpectedPointCount.Value);
                var categoryTextOk = string.IsNullOrWhiteSpace(expectedCategoryText)
                    || chartDocument.Descendants(c + "v").Any(item =>
                        string.Equals(NormalizePlainText(item.Value), expectedCategoryText, StringComparison.OrdinalIgnoreCase));

                if (categoryOk && valueOk && matchingSeriesOk && noExtraSeriesOk && seriesNamesOk && pointCountOk && categoryTextOk)
                {
                    return new SpecialConditionEvalOutcome
                    {
                        IsPassed = true,
                        Message = $"Chart {chartPath} co data range dung cau hinh."
                    };
                }
            }

            return Fail("Khong tim thay chart co data range/series/point count/category text dung cau hinh.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelChartStyle(
            ExcelChartStyleConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Chart Style (excelChartStyleConfig trong).");
            }

            if (!config.StyleId.HasValue || config.StyleId.Value <= 0)
            {
                return Fail("excelChartStyleConfig.styleId phai lon hon 0.");
            }

            var styleParts = ResolveExcelChartStyleParts(package, config.ChartSourceFile, config.StyleSourceFile);
            foreach (var stylePath in styleParts)
            {
                if (!package.TryGetXmlDocument(stylePath, out var styleDocument, out _))
                {
                    continue;
                }

                if (int.TryParse(styleDocument.Root?.Attribute("id")?.Value, out var styleId)
                    && styleId == config.StyleId.Value)
                {
                    return new SpecialConditionEvalOutcome
                    {
                        IsPassed = true,
                        Message = $"Chart style {stylePath} co id {config.StyleId.Value}."
                    };
                }
            }

            var chartParts = ResolveExcelChartParts(package, config.ChartSourceFile);
            foreach (var chartPath in chartParts)
            {
                if (!package.TryGetXmlDocument(chartPath, out var chartDocument, out _))
                {
                    continue;
                }

                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                if (chartDocument.Descendants(c + "style").Any(item =>
                    int.TryParse(item.Attribute("val")?.Value, out var styleId) && styleId == config.StyleId.Value))
                {
                    return new SpecialConditionEvalOutcome
                    {
                        IsPassed = true,
                        Message = $"Chart {chartPath} co style val {config.StyleId.Value}."
                    };
                }
            }

            return Fail($"Khong tim thay chart style id {config.StyleId.Value}.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelTextReplacement(
            ExcelTextReplacementConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Text Replacement (excelTextReplacementConfig trong).");
            }

            var oldText = NormalizePlainText(config.OldText);
            var newText = NormalizePlainText(config.NewText);
            if (string.IsNullOrWhiteSpace(oldText) || string.IsNullOrWhiteSpace(newText))
            {
                return Fail("excelTextReplacementConfig.oldText va newText khong duoc rong.");
            }

            if (!TryGetExcelTextValuesForReplacement(config, package, out var textValues, out var scopeDescription, out var scopeError))
            {
                return Fail(scopeError);
            }

            var oldCount = CountTextMatches(textValues, oldText, config.MatchWholeWord != false);
            if (config.RequireOldTextAbsent != false && oldCount > 0)
            {
                return Fail($"Van con {oldCount} lan xuat hien '{oldText}' trong {scopeDescription} chua duoc thay bang '{newText}'.");
            }

            var newCount = CountTextMatches(textValues, newText, config.MatchWholeWord != false);
            var minNewTextOccurrences = config.RequireOldTextAbsent != false
                ? 1
                : config.MinNewTextOccurrences.GetValueOrDefault(1);
            if (newCount < minNewTextOccurrences)
            {
                return Fail($"Chi tim thay {newCount} lan '{newText}' trong {scopeDescription}, can it nhat {minNewTextOccurrences} lan.");
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Da thay '{oldText}' bang '{newText}' dung yeu cau trong {scopeDescription}."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelPrintTitles(
            ExcelPrintTitlesConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Print Titles (excelPrintTitlesConfig trong).");
            }

            var worksheetName = config.WorksheetName?.Trim();
            var expectedRows = NormalizeExcelPrintRows(config.ExpectedRows);
            if (string.IsNullOrWhiteSpace(worksheetName) || string.IsNullOrWhiteSpace(expectedRows))
            {
                return Fail("excelPrintTitlesConfig phai co worksheetName va expectedRows, vi du Costs / 1:3.");
            }

            if (!package.TryGetXmlDocument("xl/workbook.xml", out var workbookDocument, out var workbookError))
            {
                return Fail(workbookError ?? "Khong doc duoc xl/workbook.xml.");
            }

            var worksheets = GetExcelWorksheets(package);
            var worksheetIndex = worksheets.FindIndex(sheet => string.Equals(sheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheetIndex < 0)
            {
                return Fail($"Khong tim thay worksheet '{worksheetName}'.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var printTitleDefinedNames = workbookDocument
                .Descendants(x + "definedName")
                .Where(item => string.Equals(item.Attribute("name")?.Value, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var matchingDefinedName = printTitleDefinedNames
                .FirstOrDefault(item =>
                {
                    var localSheetIdText = item.Attribute("localSheetId")?.Value;
                    var localSheetMatches = int.TryParse(localSheetIdText, out var localSheetId)
                        ? localSheetId == worksheetIndex
                        : item.Value.Contains(worksheetName, StringComparison.OrdinalIgnoreCase);

                    return localSheetMatches
                        && string.Equals(NormalizeExcelPrintRows(item.Value), expectedRows, StringComparison.OrdinalIgnoreCase);
                });

            if (matchingDefinedName == null)
            {
                if (printTitleDefinedNames.Count == 0)
                {
                    return Fail($"Workbook chua luu Print Titles (_xlnm.Print_Titles). Hay bam OK trong Page Setup va Save workbook truoc khi cham lai. Can {worksheetName}!{expectedRows}.");
                }

                var foundDefinitions = string.Join(
                    "; ",
                    printTitleDefinedNames.Select(item =>
                    {
                        var localSheetIdText = item.Attribute("localSheetId")?.Value;
                        var sheetName = int.TryParse(localSheetIdText, out var localSheetId)
                            && localSheetId >= 0
                            && localSheetId < worksheets.Count
                                ? worksheets[localSheetId].Name
                                : "(khong ro sheet)";
                        var rows = NormalizeExcelPrintRows(item.Value);
                        return $"{sheetName}!{(string.IsNullOrWhiteSpace(rows) ? item.Value : rows)}";
                    }));

                return Fail($"Worksheet {worksheetName} chua lap Print Titles lap lai hang {expectedRows}. Workbook hien co: {foundDefinitions}.");
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Worksheet {worksheetName} da lap Print Titles hang {expectedRows}."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelNumberFormat(
            ExcelNumberFormatConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Number Format (excelNumberFormatConfig trong).");
            }

            var expectedRange = NormalizeExcelRangeOrColumnRange(config.Range);
            if (string.IsNullOrWhiteSpace(expectedRange))
            {
                return Fail("excelNumberFormatConfig.range khong hop le. Vi du hop le: B:E hoac B4:E20.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return Fail(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
            {
                return Fail(worksheetDocumentError ?? $"Khong tim thay {worksheetPath} trong file hoc sinh.");
            }

            if (!package.TryGetXmlDocument("xl/styles.xml", out var stylesDocument, out var stylesDocumentError))
            {
                return Fail(stylesDocumentError ?? "Khong doc duoc xl/styles.xml.");
            }

            var category = string.IsNullOrWhiteSpace(config.Category) ? null : config.Category.Trim().ToLowerInvariant();
            var allowedNumberFormatIds = config.AllowedNumberFormatIds.Count > 0
                ? config.AllowedNumberFormatIds.ToHashSet()
                : new HashSet<int> { 1, 2, 3, 4 };
            var styleNumberFormats = GetExcelStyleNumberFormats(stylesDocument);
            var columnStyles = GetExcelColumnStyles(worksheetDocument);
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var numericCells = worksheetDocument
                .Descendants(x + "c")
                .Where(cell => IsNumericExcelCell(cell) && IsCellInExcelRange(cell.Attribute("r")?.Value, expectedRange))
                .ToList();

            if (numericCells.Count == 0)
            {
                return Fail($"Khong tim thay o du lieu so nao trong range {expectedRange} tren {worksheetPath}.");
            }

            var failedCells = new List<string>();
            foreach (var cell in numericCells)
            {
                var address = NormalizeExcelCellAddress(cell.Attribute("r")?.Value);
                var styleId = GetEffectiveExcelStyleId(cell, address, columnStyles);
                if (!styleId.HasValue
                    || !styleNumberFormats.TryGetValue(styleId.Value, out var numberFormat)
                    || !ExcelNumberFormatMatches(numberFormat, category, allowedNumberFormatIds, config.DecimalPlaces, config.Symbol, config.RequireThousandsSeparator == true))
                {
                    failedCells.Add(address);
                }
            }

            if (failedCells.Count > 0)
            {
                return Fail($"Cac o so trong {expectedRange} chua dung dinh dang {DescribeExcelNumberFormatRequirement(category, allowedNumberFormatIds, config.DecimalPlaces, config.Symbol, config.RequireThousandsSeparator == true)}: {string.Join(", ", failedCells.Take(12))}{(failedCells.Count > 12 ? ", ..." : string.Empty)}.");
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Tat ca {numericCells.Count} o so trong {expectedRange} tren {worksheetPath} dung dinh dang {DescribeExcelNumberFormatRequirement(category, allowedNumberFormatIds, config.DecimalPlaces, config.Symbol, config.RequireThousandsSeparator == true)}."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelChartLegend(
            ExcelChartLegendConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Chart Legend (excelChartLegendConfig trong).");
            }

            var expectedPosition = string.IsNullOrWhiteSpace(config.Position)
                ? "t"
                : config.Position.Trim();
            var chartParts = ResolveExcelChartParts(package, config.ChartSourceFile);
            if (chartParts.Count == 0)
            {
                return Fail("Khong tim thay chart XML can kiem tra trong workbook.");
            }

            foreach (var chartPath in chartParts)
            {
                if (!package.TryGetXmlDocument(chartPath, out var chartDocument, out _))
                {
                    continue;
                }

                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                var actualPosition = chartDocument
                    .Descendants(c + "legend")
                    .Elements(c + "legendPos")
                    .Select(item => item.Attribute("val")?.Value?.Trim())
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

                if (string.Equals(actualPosition, expectedPosition, StringComparison.OrdinalIgnoreCase))
                {
                    return new SpecialConditionEvalOutcome
                    {
                        IsPassed = true,
                        Message = $"Chart {chartPath} co legend position '{expectedPosition}'."
                    };
                }
            }

            return Fail($"Khong tim thay chart co legend position '{expectedPosition}'.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelTableName(
            ExcelTableNameConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Table Name (excelTableNameConfig trong).");
            }

            var expectedName = config.ExpectedName?.Trim();
            if (string.IsNullOrWhiteSpace(expectedName))
            {
                return Fail("excelTableNameConfig.expectedName khong duoc rong.");
            }

            var tableParts = ResolveExcelTableParts(package, config.WorksheetName, config.SourceFile);
            if (tableParts.Count == 0)
            {
                return Fail("Khong tim thay table XML can kiem tra trong workbook.");
            }

            var matchingTables = tableParts
                .Select(path => new { Path = path, Document = TryParsePackageXml(package, path) })
                .Where(item => item.Document != null)
                .Select(item => new
                {
                    item.Path,
                    Table = item.Document!.Root,
                    Name = item.Document!.Root?.Attribute("name")?.Value,
                    DisplayName = item.Document!.Root?.Attribute("displayName")?.Value
                })
                .ToList();

            var expectedMatch = matchingTables.FirstOrDefault(item =>
                string.Equals(item.Name, expectedName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(item.DisplayName, expectedName, StringComparison.OrdinalIgnoreCase));

            if (expectedMatch == null)
            {
                return Fail($"Khong tim thay table co ten '{expectedName}'.");
            }

            if (config.RequireOriginalNameAbsent != false && !string.IsNullOrWhiteSpace(config.OriginalName))
            {
                var oldNameStillExists = matchingTables.Any(item =>
                    string.Equals(item.Name, config.OriginalName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(item.DisplayName, config.OriginalName, StringComparison.OrdinalIgnoreCase));

                if (oldNameStillExists)
                {
                    return Fail($"Van con table ten cu '{config.OriginalName}'.");
                }
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Table '{expectedName}' ton tai trong {expectedMatch.Path}."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelWorksheetPageSetup(
            ExcelWorksheetPageSetupConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Worksheet Page Setup (excelWorksheetPageSetupConfig trong).");
            }

            var expectedOrientation = string.IsNullOrWhiteSpace(config.Orientation)
                ? "landscape"
                : config.Orientation.Trim();

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return Fail(worksheetError);
            }

            var document = TryParsePackageXml(package, worksheetPath);
            if (document?.Root == null)
            {
                return Fail($"Khong the doc worksheet XML {worksheetPath}.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var pageSetup = document.Root.Element(x + "pageSetup");
            var actualOrientation = pageSetup?.Attribute("orientation")?.Value;

            if (!string.Equals(actualOrientation, expectedOrientation, StringComparison.OrdinalIgnoreCase))
            {
                return Fail($"Worksheet {worksheetPath} co orientation '{actualOrientation ?? "(rong)"}', can '{expectedOrientation}'.");
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Worksheet {worksheetPath} co orientation '{expectedOrientation}'."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelClearCellFormatting(
            ExcelClearCellFormattingConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Clear Cell Formatting (excelClearCellFormattingConfig trong).");
            }

            if (!TryParseExcelRange(config.Range, out var startColumn, out var startRow, out var endColumn, out var endRow))
            {
                return Fail("excelClearCellFormattingConfig.range khong hop le. Vi du hop le: A4:D4.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return Fail(worksheetError);
            }

            var document = TryParsePackageXml(package, worksheetPath);
            if (document?.Root == null)
            {
                return Fail($"Khong the doc worksheet XML {worksheetPath}.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var defaultStyleId = config.DefaultStyleId ?? 0;
            var cellsByAddress = document.Descendants(x + "c")
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Attribute("r")?.Value))
                .ToDictionary(cell => cell.Attribute("r")!.Value.ToUpperInvariant(), cell => cell, StringComparer.OrdinalIgnoreCase);

            for (var row = startRow; row <= endRow; row++)
            {
                for (var column = startColumn; column <= endColumn; column++)
                {
                    var address = $"{GetExcelColumnName(column)}{row}";
                    if (!cellsByAddress.TryGetValue(address, out var cell))
                    {
                        continue;
                    }

                    var styleText = cell.Attribute("s")?.Value;
                    if (string.IsNullOrWhiteSpace(styleText))
                    {
                        continue;
                    }

                    if (!int.TryParse(styleText, out var styleId) || styleId != defaultStyleId)
                    {
                        return Fail($"O {address} van co style id '{styleText}', can clear formatting ve style {defaultStyleId}.");
                    }
                }
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = $"Range {config.Range} tren {worksheetPath} da clear formatting."
            };
        }

        private static SpecialConditionEvalOutcome EvaluateExcelDataModelImport(
            ExcelDataModelImportConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Data Model Import (excelDataModelImportConfig trong).");
            }

            if (config.RequireConnection != false)
            {
                if (!package.XmlParts.TryGetValue("xl/connections.xml", out var connectionsXml))
                {
                    return Fail("Khong tim thay xl/connections.xml de kiem tra connection import.");
                }

                var connectionHaystack = NormalizePlainText(connectionsXml);
                if (!string.IsNullOrWhiteSpace(config.ExpectedConnectionName)
                    && !connectionHaystack.Contains(config.ExpectedConnectionName.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return Fail($"Khong tim thay connection ten '{config.ExpectedConnectionName}'.");
                }
            }

            if (!string.IsNullOrWhiteSpace(config.SourceFileName)
                && !ExcelPackageContainsText(package, config.SourceFileName.Trim())
                && !ExcelPackageContainsText(package, Path.GetFileNameWithoutExtension(config.SourceFileName.Trim())))
            {
                return Fail($"Khong tim thay dau hieu file nguon '{config.SourceFileName}' trong workbook.");
            }

            if (config.RequireImportedWorksheet != false)
            {
                var expectedWorksheetName = string.IsNullOrWhiteSpace(config.ExpectedWorksheetName)
                    ? Path.GetFileNameWithoutExtension(config.SourceFileName ?? config.ExpectedConnectionName ?? string.Empty)
                    : config.ExpectedWorksheetName.Trim();

                if (string.IsNullOrWhiteSpace(expectedWorksheetName))
                {
                    return Fail("Can cau hinh expectedWorksheetName hoac sourceFileName de kiem tra worksheet import.");
                }

                var worksheets = GetExcelWorksheets(package);
                var hasImportedWorksheet = worksheets.Any(sheet =>
                    sheet.Name.Contains(expectedWorksheetName, StringComparison.OrdinalIgnoreCase));

                if (!hasImportedWorksheet)
                {
                    return Fail($"Khong tim thay worksheet import co ten chua '{expectedWorksheetName}'.");
                }
            }

            if (config.RequireQueryTable != false && !HasExcelQueryTable(package))
            {
                return Fail("Khong tim thay query table duoc tao tu thao tac import.");
            }

            if (config.RequireDataModel != false)
            {
                var hasModelPart = package.PartNames.Any(path =>
                    path.StartsWith("xl/model/", StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith("xl/model/item", StringComparison.OrdinalIgnoreCase))
                    || package.PartNames.Any(path =>
                        path.Contains("datamodel", StringComparison.OrdinalIgnoreCase)
                        || path.Contains("dataModel", StringComparison.Ordinal))
                    || (package.XmlParts.TryGetValue("xl/connections.xml", out var connectionsXml)
                        && ContainsAnyText(
                            connectionsXml,
                            "modelConnection",
                            "worksheetDataModel",
                            "Data Model",
                            "ModelConnection"));

                if (!hasModelPart)
                {
                    return Fail("Khong tim thay dau hieu Data Model trong workbook.");
                }
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = "Workbook co connection import va dau hieu Data Model dung cau hinh."
            };
        }

        private static bool HasExcelQueryTable(OfficePackage package)
        {
            if (package.PartNames.Any(path => path.StartsWith("xl/queryTables/", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return package.XmlParts
                .Where(part => part.Key.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase))
                .Select(part => TryParsePackageXml(package, part.Key)?.Root)
                .Where(root => root != null)
                .Any(root => string.Equals(root!.Attribute("tableType")?.Value, "queryTable", StringComparison.OrdinalIgnoreCase));
        }

        private static bool ExcelPackageContainsText(OfficePackage package, string expectedText)
        {
            if (string.IsNullOrWhiteSpace(expectedText))
            {
                return true;
            }

            var normalizedExpected = NormalizePlainText(expectedText);
            var partNameHaystack = package.GetNormalizedPartNameHaystack();
            if (partNameHaystack.Contains(normalizedExpected, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return package.GetNormalizedXmlHaystack()
                .Contains(normalizedExpected, StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsAnyText(string value, params string[] expectedTexts)
        {
            return expectedTexts.Any(expected =>
                !string.IsNullOrWhiteSpace(expected)
                && value.Contains(expected, StringComparison.OrdinalIgnoreCase));
        }

        private static SpecialConditionEvalOutcome EvaluateExcelCompatibilityReport(
            ExcelCompatibilityReportConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Excel Compatibility Report (excelCompatibilityReportConfig trong).");
            }

            var worksheets = GetExcelWorksheets(package);
            if (config.RequireNewWorksheet != false && worksheets.Count < 2)
            {
                return Fail("Workbook chua co worksheet moi de chua ket qua Compatibility Checker.");
            }

            if (!string.IsNullOrWhiteSpace(config.WorksheetName)
                && worksheets.All(sheet => !string.Equals(sheet.Name, config.WorksheetName.Trim(), StringComparison.OrdinalIgnoreCase)))
            {
                return Fail($"Khong tim thay worksheet '{config.WorksheetName}'.");
            }

            var expectedTexts = config.ExpectedTexts?
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(NormalizePlainText)
                .ToList() ?? new List<string>();

            if (expectedTexts.Count > 0)
            {
                var textHaystack = BuildExcelTextHaystack(package);

                var missing = expectedTexts
                    .Where(text => !textHaystack.Contains(text, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (missing.Count > 0)
                {
                    return Fail($"Thieu {missing.Count}/{expectedTexts.Count} doan text compatibility report.");
                }
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = "Workbook co worksheet/ket qua Compatibility Checker dung cau hinh."
            };
        }

        private static string BuildExcelTextHaystack(OfficePackage package)
        {
            var textValues = new List<string>();

            foreach (var sourceFile in package.XmlParts.Keys.Where(path =>
                path.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)))
            {
                var document = TryParsePackageXml(package, sourceFile);
                if (document?.Root == null)
                {
                    continue;
                }

                textValues.AddRange(document
                    .Descendants()
                    .Where(element => element.Name.LocalName is "t" or "v")
                    .Select(element => element.Value)
                    .Where(value => !string.IsNullOrWhiteSpace(value)));
            }

            return NormalizePlainText(string.Join(" ", textValues));
        }

        private static List<string> ResolveExcelTableParts(OfficePackage package, string? worksheetName, string? sourceFile)
        {
            if (!string.IsNullOrWhiteSpace(sourceFile))
            {
                var normalizedSource = NormalizeSourceFile(sourceFile);
                return package.XmlParts.ContainsKey(normalizedSource)
                    ? new List<string> { normalizedSource }
                    : new List<string>();
            }

            if (!string.IsNullOrWhiteSpace(worksheetName)
                && TryResolveExcelWorksheet(package, worksheetName, null, out var worksheetPath, out _))
            {
                return GetTablePartsForWorksheet(package, worksheetPath);
            }

            return package.XmlParts.Keys
                .Where(path => path.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase)
                    && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool TryResolveExcelWorksheet(
            OfficePackage package,
            string? worksheetName,
            string? sourceFile,
            out string worksheetPath,
            out string errorMessage)
        {
            worksheetPath = string.Empty;
            errorMessage = string.Empty;

            if (!string.IsNullOrWhiteSpace(sourceFile))
            {
                var normalizedSource = NormalizeSourceFile(sourceFile);
                if (package.XmlParts.ContainsKey(normalizedSource))
                {
                    worksheetPath = normalizedSource;
                    return true;
                }

                errorMessage = $"Khong tim thay worksheet XML {normalizedSource}.";
                return false;
            }

            var worksheets = GetExcelWorksheets(package);
            if (!string.IsNullOrWhiteSpace(worksheetName))
            {
                var match = worksheets.FirstOrDefault(sheet =>
                    string.Equals(sheet.Name, worksheetName.Trim(), StringComparison.OrdinalIgnoreCase));

                if (match != null)
                {
                    worksheetPath = match.SourceFile;
                    return true;
                }

                errorMessage = $"Khong tim thay worksheet '{worksheetName}'.";
                return false;
            }

            if (worksheets.Count == 1)
            {
                worksheetPath = worksheets[0].SourceFile;
                return true;
            }

            errorMessage = "Can cau hinh worksheetName hoac sourceFile vi workbook co nhieu worksheet.";
            return false;
        }

        private static List<ExcelWorksheetInfo> GetExcelWorksheets(OfficePackage package)
        {
            if (!package.TryGetXmlDocument("xl/workbook.xml", out var workbook, out _)
                || !package.TryGetRelationships("xl/_rels/workbook.xml.rels", out var relationships, out _))
            {
                return new List<ExcelWorksheetInfo>();
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            return workbook.Descendants(x + "sheet")
                .Select(sheet =>
                {
                    var relationshipId = sheet.Attribute(r + "id")?.Value ?? string.Empty;
                    relationships.TryGetValue(relationshipId, out var target);
                    return new ExcelWorksheetInfo
                    {
                        Name = sheet.Attribute("name")?.Value ?? string.Empty,
                        RelationshipId = relationshipId,
                        SourceFile = NormalizeRelationshipTarget("xl", target ?? string.Empty)
                    };
                })
                .Where(sheet => !string.IsNullOrWhiteSpace(sheet.SourceFile))
                .ToList();
        }

        private static List<string> GetTablePartsForWorksheet(OfficePackage package, string worksheetPath)
        {
            var relsPath = GetRelationshipPartPath(worksheetPath);
            if (!package.TryGetXmlDocument(worksheetPath, out var worksheet, out _)
                || !package.TryGetRelationships(relsPath, out var relationships, out _))
            {
                return new List<string>();
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var baseFolder = GetPackageFolder(worksheetPath);

            return worksheet.Descendants(x + "tablePart")
                .Select(tablePart => tablePart.Attribute(r + "id")?.Value)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => relationships.TryGetValue(id!, out var target)
                    ? NormalizeRelationshipTarget(baseFolder, target)
                    : string.Empty)
                .Where(path => !string.IsNullOrWhiteSpace(path) && package.XmlParts.ContainsKey(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> ResolveExcelChartParts(OfficePackage package, string? chartSourceFile)
        {
            if (!string.IsNullOrWhiteSpace(chartSourceFile))
            {
                var normalizedSource = NormalizeSourceFile(chartSourceFile);
                return package.XmlParts.ContainsKey(normalizedSource)
                    ? new List<string> { normalizedSource }
                    : new List<string>();
            }

            return package.XmlParts.Keys
                .Where(path => path.StartsWith("xl/charts/chart", StringComparison.OrdinalIgnoreCase)
                    && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> ResolveExcelChartStyleParts(
            OfficePackage package,
            string? chartSourceFile,
            string? styleSourceFile)
        {
            if (!string.IsNullOrWhiteSpace(styleSourceFile))
            {
                var normalizedSource = NormalizeSourceFile(styleSourceFile);
                return package.XmlParts.ContainsKey(normalizedSource)
                    ? new List<string> { normalizedSource }
                    : new List<string>();
            }

            var styleParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var chartParts = ResolveExcelChartParts(package, chartSourceFile);

            foreach (var chartPath in chartParts)
            {
                var relsPath = GetRelationshipPartPath(chartPath);
                if (!package.TryGetRelationships(relsPath, out var relationships, out _))
                {
                    continue;
                }

                var baseFolder = GetPackageFolder(chartPath);
                foreach (var target in relationships.Values)
                {
                    var normalizedTarget = NormalizeRelationshipTarget(baseFolder, target);
                    if (normalizedTarget.StartsWith("xl/charts/style", StringComparison.OrdinalIgnoreCase)
                        && normalizedTarget.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        && package.XmlParts.ContainsKey(normalizedTarget))
                    {
                        styleParts.Add(normalizedTarget);
                    }
                }
            }

            foreach (var path in package.XmlParts.Keys.Where(path =>
                path.StartsWith("xl/charts/style", StringComparison.OrdinalIgnoreCase)
                && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            {
                styleParts.Add(path);
            }

            return styleParts.OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static bool RangeListContains(string? sqref, string expectedRange)
        {
            if (string.IsNullOrWhiteSpace(sqref) || string.IsNullOrWhiteSpace(expectedRange))
            {
                return false;
            }

            return sqref
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Any(range => string.Equals(
                    NormalizeExcelRangeAddress(range),
                    NormalizeExcelRangeAddress(expectedRange),
                    StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeExcelCellAddress(string? cell)
        {
            if (string.IsNullOrWhiteSpace(cell))
            {
                return string.Empty;
            }

            var normalized = cell.Trim().Replace("$", string.Empty).ToUpperInvariant();
            return Regex.IsMatch(normalized, "^[A-Z]{1,3}[1-9][0-9]*$", RegexOptions.CultureInvariant)
                ? normalized
                : string.Empty;
        }

        private static string NormalizeExcelRangeAddress(string? range)
        {
            if (string.IsNullOrWhiteSpace(range))
            {
                return string.Empty;
            }

            var normalized = range.Trim().Replace("$", string.Empty).ToUpperInvariant();
            if (!TryParseExcelRange(normalized, out var startColumn, out var startRow, out var endColumn, out var endRow))
            {
                return string.Empty;
            }

            var start = $"{GetExcelColumnName(startColumn)}{startRow}";
            var end = $"{GetExcelColumnName(endColumn)}{endRow}";
            return string.Equals(start, end, StringComparison.OrdinalIgnoreCase)
                ? start
                : $"{start}:{end}";
        }

        private static string NormalizeExcelFormulaReference(string? reference)
        {
            if (string.IsNullOrWhiteSpace(reference))
            {
                return string.Empty;
            }

            var normalized = reference.Trim().Replace("'", string.Empty).Replace("$", string.Empty);
            normalized = Regex.Replace(normalized, @"\s+", string.Empty);
            return normalized.ToUpperInvariant();
        }

        private static string NormalizeExcelPrintRows(string? rows)
        {
            if (string.IsNullOrWhiteSpace(rows))
            {
                return string.Empty;
            }

            var normalized = rows.Trim().Replace("'", string.Empty).Replace("$", string.Empty);
            var match = Regex.Match(normalized, @"(?<![A-Z])(\d+):(\d+)(?![A-Z0-9])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return string.Empty;
            }

            var start = int.Parse(match.Groups[1].Value);
            var end = int.Parse(match.Groups[2].Value);
            if (start <= 0 || end <= 0)
            {
                return string.Empty;
            }

            if (start > end)
            {
                (start, end) = (end, start);
            }

            return $"{start}:{end}";
        }

        private static string NormalizeExcelRangeOrColumnRange(string? range)
        {
            var normalizedRange = NormalizeExcelRangeAddress(range);
            if (!string.IsNullOrWhiteSpace(normalizedRange))
            {
                return normalizedRange;
            }

            if (!TryParseExcelColumnRange(range, out var startColumn, out var endColumn))
            {
                return string.Empty;
            }

            return $"{GetExcelColumnName(startColumn)}:{GetExcelColumnName(endColumn)}";
        }

        private static bool TryParseExcelColumnRange(string? range, out int startColumn, out int endColumn)
        {
            startColumn = 0;
            endColumn = 0;
            if (string.IsNullOrWhiteSpace(range))
            {
                return false;
            }

            var match = Regex.Match(
                range.Trim().Replace("$", string.Empty),
                @"^([A-Z]{1,3})(?::([A-Z]{1,3}))?$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return false;
            }

            startColumn = GetExcelColumnNumber(match.Groups[1].Value);
            endColumn = match.Groups[2].Success ? GetExcelColumnNumber(match.Groups[2].Value) : startColumn;
            if (startColumn <= 0 || endColumn <= 0)
            {
                return false;
            }

            if (startColumn > endColumn)
            {
                (startColumn, endColumn) = (endColumn, startColumn);
            }

            return true;
        }

        private static bool IsCellInExcelRange(string? cellAddress, string range)
        {
            var normalizedCell = NormalizeExcelCellAddress(cellAddress);
            if (string.IsNullOrWhiteSpace(normalizedCell) || string.IsNullOrWhiteSpace(range))
            {
                return false;
            }

            var match = Regex.Match(normalizedCell, "^([A-Z]{1,3})([1-9][0-9]*)$", RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return false;
            }

            var column = GetExcelColumnNumber(match.Groups[1].Value);
            var row = int.Parse(match.Groups[2].Value);

            if (TryParseExcelColumnRange(range, out var startColumn, out var endColumn))
            {
                return column >= startColumn && column <= endColumn;
            }

            return TryParseExcelRange(range, out startColumn, out var startRow, out endColumn, out var endRow)
                && column >= startColumn
                && column <= endColumn
                && row >= startRow
                && row <= endRow;
        }

        private static List<string> GetExcelTextValues(OfficePackage package)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            return package.XmlParts
                .Where(part => string.Equals(part.Key, "xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)
                    || part.Key.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase))
                .Select(part => TryParsePackageXml(package, part.Key))
                .Where(document => document != null)
                .SelectMany(document => document!.Descendants(x + "t"))
                .Select(item => NormalizePlainText(item.Value))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList();
        }

        private static bool TryGetExcelTextValuesForReplacement(
            ExcelTextReplacementConfig config,
            OfficePackage package,
            out List<string> values,
            out string scopeDescription,
            out string error)
        {
            values = new List<string>();
            error = string.Empty;

            if (string.IsNullOrWhiteSpace(config.WorksheetName) && string.IsNullOrWhiteSpace(config.SourceFile))
            {
                scopeDescription = "toan bo workbook";
                values = GetExcelTextValues(package);
                return true;
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                scopeDescription = "worksheet";
                error = worksheetError;
                return false;
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
            {
                scopeDescription = worksheetPath;
                error = worksheetDocumentError ?? $"Khong tim thay {worksheetPath} trong file hoc sinh.";
                return false;
            }

            scopeDescription = string.IsNullOrWhiteSpace(config.WorksheetName)
                ? worksheetPath
                : $"worksheet {config.WorksheetName.Trim()}";
            values = GetExcelTextValuesFromWorksheet(package, worksheetDocument);
            return true;
        }

        private static List<string> GetExcelTextValuesFromWorksheet(
            OfficePackage package,
            XDocument worksheetDocument)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sharedStrings = GetExcelSharedStrings(package);
            var values = new List<string>();

            foreach (var cell in worksheetDocument.Descendants(x + "c"))
            {
                var type = cell.Attribute("t")?.Value;
                if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase))
                {
                    var sharedStringIndexText = cell.Element(x + "v")?.Value;
                    if (int.TryParse(sharedStringIndexText, out var sharedStringIndex)
                        && sharedStringIndex >= 0
                        && sharedStringIndex < sharedStrings.Count)
                    {
                        values.Add(sharedStrings[sharedStringIndex]);
                    }

                    continue;
                }

                if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(type, "str", StringComparison.OrdinalIgnoreCase))
                {
                    values.AddRange(cell.Descendants(x + "t").Select(item => NormalizePlainText(item.Value)));
                    var formulaStringValue = cell.Element(x + "v")?.Value;
                    if (!string.IsNullOrWhiteSpace(formulaStringValue))
                    {
                        values.Add(NormalizePlainText(formulaStringValue));
                    }
                }
            }

            return values
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList();
        }

        private static List<string> GetExcelSharedStrings(OfficePackage package)
        {
            if (!package.TryGetXmlDocument("xl/sharedStrings.xml", out var sharedStringsDocument, out _))
            {
                return new List<string>();
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            return sharedStringsDocument
                .Descendants(x + "si")
                .Select(item => NormalizePlainText(string.Concat(item.Descendants(x + "t").Select(text => text.Value))))
                .ToList();
        }

        private static int CountTextMatches(IEnumerable<string> values, string expectedText, bool wholeWord)
        {
            if (string.IsNullOrWhiteSpace(expectedText))
            {
                return 0;
            }

            if (!wholeWord)
            {
                return values.Sum(value => Regex.Matches(
                    value,
                    Regex.Escape(expectedText),
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count);
            }

            var pattern = $@"(?<![\p{{L}}\p{{N}}_]){Regex.Escape(expectedText)}(?![\p{{L}}\p{{N}}_])";
            return values.Sum(value => Regex.Matches(
                value,
                pattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count);
        }

        private static bool IsNumericExcelCell(XElement cell)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var cellType = cell.Attribute("t")?.Value;
            return (string.IsNullOrWhiteSpace(cellType)
                    || string.Equals(cellType, "n", StringComparison.OrdinalIgnoreCase))
                && cell.Element(x + "v") != null;
        }

        private sealed class ExcelNumberFormatInfo
        {
            public int NumberFormatId { get; init; }
            public string FormatCode { get; init; } = string.Empty;
        }

        private static Dictionary<int, ExcelNumberFormatInfo> GetExcelStyleNumberFormats(XDocument stylesDocument)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var customFormats = stylesDocument
                .Root?
                .Element(x + "numFmts")?
                .Elements(x + "numFmt")
                .Where(element => int.TryParse(element.Attribute("numFmtId")?.Value, out _))
                .ToDictionary(
                    element => int.Parse(element.Attribute("numFmtId")!.Value),
                    element => element.Attribute("formatCode")?.Value ?? string.Empty)
                ?? new Dictionary<int, string>();

            var cellXfs = stylesDocument
                .Root?
                .Element(x + "cellXfs")?
                .Elements(x + "xf")
                .ToList() ?? new List<XElement>();

            var result = new Dictionary<int, ExcelNumberFormatInfo>();
            for (var index = 0; index < cellXfs.Count; index++)
            {
                if (int.TryParse(cellXfs[index].Attribute("numFmtId")?.Value, out var numberFormatId))
                {
                    var formatCode = customFormats.TryGetValue(numberFormatId, out var customFormatCode)
                        ? customFormatCode
                        : BuiltInExcelNumberFormats.GetValueOrDefault(numberFormatId, string.Empty);
                    result[index] = new ExcelNumberFormatInfo
                    {
                        NumberFormatId = numberFormatId,
                        FormatCode = formatCode
                    };
                }
            }

            return result;
        }

        private static bool ExcelNumberFormatMatches(
            ExcelNumberFormatInfo numberFormat,
            string? category,
            IReadOnlySet<int> allowedNumberFormatIds,
            int? decimalPlaces,
            string? symbol,
            bool requireThousandsSeparator)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return allowedNumberFormatIds.Contains(numberFormat.NumberFormatId);
            }

            var formatCode = numberFormat.FormatCode;
            if (string.IsNullOrWhiteSpace(formatCode))
            {
                formatCode = BuiltInExcelNumberFormats.GetValueOrDefault(numberFormat.NumberFormatId, string.Empty);
            }

            if (!ExcelFormatMatchesCategory(numberFormat.NumberFormatId, formatCode, category))
            {
                return false;
            }

            if (decimalPlaces.HasValue && CountExcelFormatDecimalPlaces(formatCode) != decimalPlaces.Value)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(symbol)
                && !formatCode.Contains(symbol.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return !requireThousandsSeparator || GetExcelNumberFormatFirstSection(formatCode).Contains(',');
        }

        private static bool ExcelFormatMatchesCategory(int numberFormatId, string formatCode, string category)
        {
            var firstSection = GetExcelNumberFormatFirstSection(formatCode);
            var comparable = RemoveExcelFormatLiterals(firstSection).ToLowerInvariant();
            return category.Trim().ToLowerInvariant() switch
            {
                "general" => numberFormatId == 0 || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase),
                "number" => IsExcelNumberFormat(numberFormatId, firstSection, comparable),
                "currency" => !IsExcelPercentFormat(numberFormatId, comparable)
                    && !IsExcelDateOrTimeFormat(numberFormatId, comparable)
                    && ContainsExcelCurrencyMarker(firstSection),
                "accounting" => !IsExcelPercentFormat(numberFormatId, comparable)
                    && !IsExcelDateOrTimeFormat(numberFormatId, comparable)
                    && ContainsExcelCurrencyMarker(firstSection)
                    && (firstSection.Contains("_", StringComparison.Ordinal) || firstSection.Contains("*", StringComparison.Ordinal)),
                "percentage" => IsExcelPercentFormat(numberFormatId, comparable),
                "date" => IsExcelDateOrTimeFormat(numberFormatId, comparable) && Regex.IsMatch(comparable, "[dmy]", RegexOptions.CultureInvariant),
                "time" => IsExcelDateOrTimeFormat(numberFormatId, comparable) && Regex.IsMatch(comparable, "[hs]", RegexOptions.CultureInvariant),
                "custom" => true,
                _ => false
            };
        }

        private static bool IsExcelNumberFormat(int numberFormatId, string rawFormatCode, string comparableFormatCode)
        {
            return numberFormatId is 1 or 2 or 3 or 4 or 37 or 38 or 39 or 40
                || (!string.Equals(comparableFormatCode, "general", StringComparison.OrdinalIgnoreCase)
                    && !IsExcelPercentFormat(numberFormatId, comparableFormatCode)
                    && !IsExcelDateOrTimeFormat(numberFormatId, comparableFormatCode)
                    && !ContainsExcelCurrencyMarker(rawFormatCode)
                    && Regex.IsMatch(comparableFormatCode, "[0#?]", RegexOptions.CultureInvariant));
        }

        private static bool IsExcelPercentFormat(int numberFormatId, string comparableFormatCode)
        {
            return numberFormatId is 9 or 10 || comparableFormatCode.Contains('%');
        }

        private static bool IsExcelDateOrTimeFormat(int numberFormatId, string comparableFormatCode)
        {
            return numberFormatId is >= 14 and <= 22
                || Regex.IsMatch(comparableFormatCode, @"(^|[^\\])([dmyhs])", RegexOptions.CultureInvariant);
        }

        private static bool ContainsExcelCurrencyMarker(string formatCode)
        {
            return Regex.IsMatch(formatCode, @"(\$|â‚¬|Â£|Â¥|â‚«|â‚©|â‚¹|â‚½|à¸¿|â‚±|vnd|usd|eur|gbp|jpy)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static int CountExcelFormatDecimalPlaces(string formatCode)
        {
            var firstSection = RemoveExcelFormatLiterals(GetExcelNumberFormatFirstSection(formatCode));
            var decimalIndex = firstSection.IndexOf('.');
            if (decimalIndex < 0)
            {
                return 0;
            }

            var count = 0;
            for (var index = decimalIndex + 1; index < firstSection.Length; index++)
            {
                if (firstSection[index] is '0' or '#' or '?')
                {
                    count++;
                    continue;
                }

                break;
            }

            return count;
        }

        private static string GetExcelNumberFormatFirstSection(string formatCode)
        {
            if (string.IsNullOrWhiteSpace(formatCode))
            {
                return string.Empty;
            }

            return formatCode.Split(';', 2)[0];
        }

        private static string RemoveExcelFormatLiterals(string formatCode)
        {
            if (string.IsNullOrWhiteSpace(formatCode))
            {
                return string.Empty;
            }

            var withoutQuotedText = Regex.Replace(formatCode, "\"[^\"]*\"", string.Empty, RegexOptions.CultureInvariant);
            var withoutBracketCodes = Regex.Replace(withoutQuotedText, @"\[[^\]]+\]", string.Empty, RegexOptions.CultureInvariant);
            return Regex.Replace(withoutBracketCodes, @"\\.", string.Empty, RegexOptions.CultureInvariant);
        }

        private static string DescribeExcelNumberFormatRequirement(
            string? category,
            IReadOnlySet<int> allowedNumberFormatIds,
            int? decimalPlaces,
            string? symbol,
            bool requireThousandsSeparator)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return $"numFmtId thuoc [{string.Join(", ", allowedNumberFormatIds.OrderBy(value => value))}]";
            }

            var parts = new List<string> { $"category {category}" };
            if (decimalPlaces.HasValue)
            {
                parts.Add($"{decimalPlaces.Value} decimal places");
            }
            if (!string.IsNullOrWhiteSpace(symbol))
            {
                parts.Add($"symbol {symbol}");
            }
            if (requireThousandsSeparator)
            {
                parts.Add("co thousands separator");
            }

            return string.Join(", ", parts);
        }

        private static Dictionary<int, int> GetExcelColumnStyles(XDocument worksheetDocument)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var result = new Dictionary<int, int>();
            foreach (var column in worksheetDocument.Descendants(x + "col"))
            {
                if (!int.TryParse(column.Attribute("min")?.Value, out var min)
                    || !int.TryParse(column.Attribute("max")?.Value, out var max)
                    || !int.TryParse(column.Attribute("style")?.Value, out var styleId))
                {
                    continue;
                }

                for (var index = min; index <= max; index++)
                {
                    result[index] = styleId;
                }
            }

            return result;
        }

        private static int? GetEffectiveExcelStyleId(
            XElement cell,
            string cellAddress,
            IReadOnlyDictionary<int, int> columnStyles)
        {
            if (int.TryParse(cell.Attribute("s")?.Value, out var cellStyleId))
            {
                return cellStyleId;
            }

            var match = Regex.Match(cellAddress, "^([A-Z]{1,3})[1-9][0-9]*$", RegexOptions.CultureInvariant);
            if (match.Success && columnStyles.TryGetValue(GetExcelColumnNumber(match.Groups[1].Value), out var columnStyleId))
            {
                return columnStyleId;
            }

            return null;
        }

        private static SpecialConditionEvalOutcome EvaluateExcelDefinedName(
            ExcelDefinedNameConfig? config,
            OfficePackage package)
        {
            if (config == null || string.IsNullOrWhiteSpace(config.Name) || config.ExpectedRanges.Count == 0)
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh tÃªn vÃ¹ng hoáº·c vÃ¹ng Ã´ cáº§n kiá»ƒm tra.");
            }

            if (!package.TryGetXmlDocument("xl/workbook.xml", out var workbookDocument, out var workbookError))
            {
                return FailSpecialCondition(workbookError ?? "KhÃ´ng tÃ¬m tháº¥y xl/workbook.xml trong file há»c sinh.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var definedName = workbookDocument
                .Descendants(x + "definedName")
                .FirstOrDefault(item => string.Equals(item.Attribute("name")?.Value, config.Name, StringComparison.OrdinalIgnoreCase));

            if (definedName == null)
            {
                return FailSpecialCondition($"ChÆ°a táº¡o named range '{config.Name}'.");
            }

            var actualRanges = SplitExcelDefinedNameRanges(definedName.Value);
            var missingRanges = config.ExpectedRanges
                .Where(expected => !actualRanges.Any(actual => ExcelDefinedNameRangeMatches(actual, expected)))
                .ToList();

            if (missingRanges.Count > 0)
            {
                return FailSpecialCondition($"Named range '{config.Name}' chÆ°a bao gá»“m Ä‘Ãºng vÃ¹ng: {string.Join(", ", missingRanges)}.");
            }

            if (config.RequireExactRanges == true
                && (actualRanges.Count != config.ExpectedRanges.Count
                    || actualRanges.Any(actual => !config.ExpectedRanges.Any(expected => ExcelDefinedNameRangeMatches(actual, expected)))))
            {
                return FailSpecialCondition($"Named range '{config.Name}' cÃ³ thÃªm hoáº·c thiáº¿u vÃ¹ng ngoÃ i cáº¥u hÃ¬nh yÃªu cáº§u.");
            }

            return PassSpecialCondition($"Named range '{config.Name}' Ä‘Ã£ trá» Ä‘Ãºng cÃ¡c vÃ¹ng yÃªu cáº§u.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelFormulaReferences(
            ExcelFormulaReferencesConfig? config,
            OfficePackage package)
        {
            if (config == null || string.IsNullOrWhiteSpace(config.Cell))
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh Ã´ cÃ´ng thá»©c hoáº·c cÃ¡c named range báº¯t buá»™c.");
            }

            if (config.RequiredReferences.Count == 0
                && config.RequiredFunctions.Count == 0
                && config.RequiredFormulaFragments.Count == 0
                && string.IsNullOrWhiteSpace(config.ExpectedFormula)
                && string.IsNullOrWhiteSpace(config.ExpectedValue))
            {
                return FailSpecialCondition("Chưa cấu hình tiêu chí kiểm tra công thức.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return FailSpecialCondition(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
            {
                return FailSpecialCondition(worksheetDocumentError ?? $"KhÃ´ng tÃ¬m tháº¥y {worksheetPath} trong file há»c sinh.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var cell = worksheetDocument
                .Descendants(x + "c")
                .FirstOrDefault(item => string.Equals(NormalizeExcelCellAddress(item.Attribute("r")?.Value), config.Cell, StringComparison.OrdinalIgnoreCase));

            if (cell == null)
            {
                return FailSpecialCondition($"KhÃ´ng tÃ¬m tháº¥y Ã´ {config.Cell} trong worksheet cáº§n kiá»ƒm tra.");
            }

            var formula = cell.Element(x + "f")?.Value;
            var normalizedFormula = NormalizeExcelFormulaText(formula);
            if (string.IsNullOrWhiteSpace(normalizedFormula))
            {
                return FailSpecialCondition($"Ã” {config.Cell} chÆ°a cÃ³ cÃ´ng thá»©c.");
            }

            if (!string.IsNullOrWhiteSpace(config.ExpectedFormula)
                && !string.Equals(normalizedFormula, config.ExpectedFormula, StringComparison.OrdinalIgnoreCase))
            {
                return FailSpecialCondition($"CÃ´ng thá»©c táº¡i Ã´ {config.Cell} chÆ°a khá»›p cÃ´ng thá»©c yÃªu cáº§u.");
            }

            var missingReferences = config.RequiredReferences
                .Where(reference => !ExcelFormulaContainsName(normalizedFormula, reference))
                .ToList();
            if (missingReferences.Count > 0)
            {
                return FailSpecialCondition($"CÃ´ng thá»©c táº¡i Ã´ {config.Cell} chÆ°a dÃ¹ng Ä‘á»§ named range: {string.Join(", ", missingReferences)}.");
            }

            var missingFunctions = config.RequiredFunctions
                .Where(functionName => !ExcelFormulaContainsFunction(normalizedFormula, functionName))
                .ToList();
            if (missingFunctions.Count > 0)
            {
                return FailSpecialCondition($"CÃ´ng thá»©c táº¡i Ã´ {config.Cell} chÆ°a dÃ¹ng Ä‘á»§ hÃ m báº¯t buá»™c: {string.Join(", ", missingFunctions)}.");
            }

            var normalizedFragmentFormula = NormalizeExcelFormulaFragment(formula);
            var missingFragments = config.RequiredFormulaFragments
                .Where(fragment => !normalizedFragmentFormula.Contains(fragment, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (missingFragments.Count > 0)
            {
                return FailSpecialCondition($"CÃ´ng thá»©c táº¡i Ã´ {config.Cell} chÆ°a cÃ³ Ä‘á»§ pháº§n cÃ´ng thá»©c báº¯t buá»™c.");
            }

            if (config.RequireOnlyDefinedNameReferences == true && ExcelFormulaContainsCellReference(normalizedFormula))
            {
                return FailSpecialCondition($"CÃ´ng thá»©c táº¡i Ã´ {config.Cell} Ä‘ang tham chiáº¿u trá»±c tiáº¿p Ã´/vÃ¹ng thay vÃ¬ chá»‰ dÃ¹ng named range.");
            }

            if (!string.IsNullOrWhiteSpace(config.ExpectedValue))
            {
                var actualValue = NormalizePlainText(cell.Element(x + "v")?.Value ?? string.Empty);
                if (!string.Equals(actualValue, config.ExpectedValue, StringComparison.OrdinalIgnoreCase))
                {
                    return FailSpecialCondition($"GiÃ¡ trá»‹ hiá»ƒn thá»‹/lÆ°u táº¡i Ã´ {config.Cell} chÆ°a Ä‘Ãºng.");
                }
            }

            return PassSpecialCondition($"Ã” {config.Cell} Ä‘Ã£ dÃ¹ng Ä‘Ãºng cÃ´ng thá»©c vá»›i cÃ¡c named range yÃªu cáº§u.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelNoConditionalFormatting(
            ExcelNoConditionalFormattingConfig? config,
            OfficePackage package)
        {
            if (config == null)
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh worksheet cáº§n kiá»ƒm tra conditional formatting.");
            }

            var worksheetPaths = new List<string>();
            if (config.RequireAllWorksheets == true)
            {
                worksheetPaths.AddRange(GetExcelWorksheets(package).Select(sheet => sheet.SourceFile));
            }
            else
            {
                if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
                {
                    return FailSpecialCondition(worksheetError);
                }

                worksheetPaths.Add(worksheetPath);
            }

            foreach (var worksheetPath in worksheetPaths.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
                {
                    return FailSpecialCondition(worksheetDocumentError ?? $"KhÃ´ng tÃ¬m tháº¥y {worksheetPath} trong file há»c sinh.");
                }

                var hasConditionalFormatting = worksheetDocument
                    .Descendants()
                    .Any(element => string.Equals(element.Name.LocalName, "conditionalFormatting", StringComparison.OrdinalIgnoreCase));

                if (hasConditionalFormatting)
                {
                    return FailSpecialCondition($"Worksheet {worksheetPath} váº«n cÃ²n conditional formatting.");
                }
            }

            return PassSpecialCondition("Worksheet Ä‘Ã£ Ä‘Æ°á»£c xÃ³a toÃ n bá»™ conditional formatting theo yÃªu cáº§u.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelTextRotation(
            ExcelTextRotationConfig? config,
            OfficePackage package)
        {
            if (config == null || config.ExpectedTexts.Count == 0 || config.AllowedTextRotationValues.Count == 0)
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh vÄƒn báº£n hoáº·c gÃ³c xoay cáº§n kiá»ƒm tra.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return FailSpecialCondition(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
            {
                return FailSpecialCondition(worksheetDocumentError ?? $"KhÃ´ng tÃ¬m tháº¥y {worksheetPath} trong file há»c sinh.");
            }

            if (!package.TryGetXmlDocument("xl/styles.xml", out var stylesDocument, out _))
            {
                return FailSpecialCondition("KhÃ´ng tÃ¬m tháº¥y xl/styles.xml Ä‘á»ƒ kiá»ƒm tra gÃ³c xoay chá»¯.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sharedStrings = GetExcelSharedStrings(package);
            var columnStyles = GetExcelColumnStyles(worksheetDocument);
            var textRotations = GetExcelStyleTextRotations(stylesDocument);
            var matchedTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var cell in worksheetDocument.Descendants(x + "c"))
            {
                var cellText = NormalizePlainText(GetExcelCellDisplayText(cell, sharedStrings));
                if (string.IsNullOrWhiteSpace(cellText))
                {
                    continue;
                }

                var expectedText = config.ExpectedTexts.FirstOrDefault(text => string.Equals(text, cellText, StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrWhiteSpace(expectedText))
                {
                    continue;
                }

                var cellAddress = NormalizeExcelCellAddress(cell.Attribute("r")?.Value);
                var styleId = GetEffectiveExcelStyleId(cell, cellAddress, columnStyles);
                var rotation = styleId.HasValue && textRotations.TryGetValue(styleId.Value, out var value) ? value : 0;
                if (config.AllowedTextRotationValues.Contains(rotation))
                {
                    matchedTexts.Add(expectedText);
                }
            }

            var missingTexts = config.ExpectedTexts
                .Where(text => !matchedTexts.Contains(text))
                .ToList();

            if (missingTexts.Count > 0 && config.RequireAllTexts == true)
            {
                return FailSpecialCondition($"CÃ¡c tiÃªu Ä‘á» sau chÆ°a Ä‘Æ°á»£c xoay Ä‘Ãºng Angle Counterclockwise: {string.Join(", ", missingTexts)}.");
            }

            if (matchedTexts.Count == 0)
            {
                return FailSpecialCondition("ChÆ°a tÃ¬m tháº¥y tiÃªu Ä‘á» nÃ o Ä‘Æ°á»£c xoay Ä‘Ãºng gÃ³c yÃªu cáº§u.");
            }

            return PassSpecialCondition("CÃ¡c tiÃªu Ä‘á» Ä‘Ã£ Ä‘Æ°á»£c xoay chá»¯ Ä‘Ãºng Angle Counterclockwise.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelMultiColumnSort(
            ExcelMultiColumnSortConfig? config,
            OfficePackage package)
        {
            if (config == null || config.KeyColumns.Count == 0 || config.HeaderRow <= 0)
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh worksheet, header row hoáº·c cÃ¡c cá»™t sáº¯p xáº¿p.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return FailSpecialCondition(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
            {
                return FailSpecialCondition(worksheetDocumentError ?? $"KhÃ´ng tÃ¬m tháº¥y {worksheetPath} trong file há»c sinh.");
            }

            if (!TryResolveExcelSortColumns(config, package, worksheetDocument, out var sortKeys, out var sortError))
            {
                return FailSpecialCondition(sortError);
            }

            var rows = GetExcelDataRowsForSort(config, package, worksheetDocument);
            if (rows.Count < 2)
            {
                return FailSpecialCondition("KhÃ´ng Ä‘á»§ dÃ²ng dá»¯ liá»‡u Ä‘á»ƒ kiá»ƒm tra sáº¯p xáº¿p.");
            }

            for (var index = 1; index < rows.Count; index++)
            {
                if (CompareExcelSortRows(rows[index - 1], rows[index], sortKeys) > 0)
                {
                    return FailSpecialCondition($"Dá»¯ liá»‡u chÆ°a Ä‘Æ°á»£c sáº¯p xáº¿p Ä‘Ãºng táº¡i dÃ²ng {rows[index].RowIndex}.");
                }
            }

            return PassSpecialCondition("Dá»¯ liá»‡u Ä‘Ã£ Ä‘Æ°á»£c sáº¯p xáº¿p Ä‘Ãºng theo nhiá»u cá»™t yÃªu cáº§u.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelFreezePanes(
            ExcelFreezePanesConfig? config,
            OfficePackage package)
        {
            if (config == null)
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh worksheet cáº§n kiá»ƒm tra Freeze Panes.");
            }

            if (!TryResolveExcelWorksheet(package, config.WorksheetName, config.SourceFile, out var worksheetPath, out var worksheetError))
            {
                return FailSpecialCondition(worksheetError);
            }

            if (!package.TryGetXmlDocument(worksheetPath, out var worksheetDocument, out var worksheetDocumentError))
            {
                return FailSpecialCondition(worksheetDocumentError ?? $"KhÃ´ng tÃ¬m tháº¥y {worksheetPath} trong file há»c sinh.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var pane = worksheetDocument.Descendants(x + "pane").FirstOrDefault();
            if (pane == null)
            {
                return FailSpecialCondition("Worksheet chÆ°a báº­t Freeze Panes.");
            }

            var state = pane.Attribute("state")?.Value ?? string.Empty;
            if (!string.Equals(state, "frozen", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(state, "frozenSplit", StringComparison.OrdinalIgnoreCase))
            {
                return FailSpecialCondition("Worksheet cÃ³ pane nhÆ°ng chÆ°a pháº£i Freeze Panes.");
            }

            if (!string.IsNullOrWhiteSpace(config.TopLeftCell)
                && !string.Equals(NormalizeExcelCellAddress(pane.Attribute("topLeftCell")?.Value), config.TopLeftCell, StringComparison.OrdinalIgnoreCase))
            {
                return FailSpecialCondition($"Freeze Panes chÆ°a giá»¯ Ä‘Ãºng cÃ¡c hÃ ng phÃ­a trÃªn Ã´ {config.TopLeftCell}.");
            }

            if (config.YSplit.HasValue && !ExcelDecimalAttributeEquals(pane.Attribute("ySplit")?.Value, config.YSplit.Value))
            {
                return FailSpecialCondition($"Freeze Panes chÆ°a cá»‘ Ä‘á»‹nh Ä‘Ãºng {config.YSplit.Value} hÃ ng Ä‘áº§u.");
            }

            if (config.RequireNoColumnFreeze == true)
            {
                var xSplit = ParseExcelDecimal(pane.Attribute("xSplit")?.Value) ?? 0m;
                if (xSplit != 0m)
                {
                    return FailSpecialCondition("Freeze Panes Ä‘ang cá»‘ Ä‘á»‹nh thÃªm cá»™t, trong khi task chá»‰ yÃªu cáº§u cá»‘ Ä‘á»‹nh hÃ ng.");
                }
            }
            else if (config.XSplit.HasValue && !ExcelDecimalAttributeEquals(pane.Attribute("xSplit")?.Value, config.XSplit.Value))
            {
                return FailSpecialCondition($"Freeze Panes chÆ°a cá»‘ Ä‘á»‹nh Ä‘Ãºng sá»‘ cá»™t yÃªu cáº§u: {config.XSplit.Value}.");
            }

            return PassSpecialCondition("Worksheet Ä‘Ã£ cá»‘ Ä‘á»‹nh Ä‘Ãºng hÃ ng khi cuá»™n theo chiá»u dá»c.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelDocumentProperty(
            ExcelDocumentPropertyConfig? config,
            OfficePackage package)
        {
            if (config == null || string.IsNullOrWhiteSpace(config.PropertyName) || string.IsNullOrWhiteSpace(config.ExpectedValue))
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh tÃªn thuá»™c tÃ­nh hoáº·c giÃ¡ trá»‹ cáº§n kiá»ƒm tra.");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile) ? "docProps/custom.xml" : config.SourceFile;
            if (!package.TryGetXmlDocument(sourceFile, out var propertyDocument, out var propertyError))
            {
                return FailSpecialCondition(propertyError ?? $"KhÃ´ng tÃ¬m tháº¥y {sourceFile} trong file há»c sinh.");
            }

            XNamespace custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
            var property = propertyDocument
                .Descendants(custom + "property")
                .FirstOrDefault(item => string.Equals(item.Attribute("name")?.Value, config.PropertyName, StringComparison.OrdinalIgnoreCase));

            if (property == null)
            {
                return FailSpecialCondition($"ChÆ°a thÃªm thuá»™c tÃ­nh '{config.PropertyName}' vÃ o workbook.");
            }

            var actualValue = NormalizePlainText(string.Concat(property.Elements().Select(element => element.Value)));
            if (!string.Equals(actualValue, config.ExpectedValue, StringComparison.OrdinalIgnoreCase))
            {
                return FailSpecialCondition($"Thuá»™c tÃ­nh '{config.PropertyName}' chÆ°a cÃ³ giÃ¡ trá»‹ '{config.ExpectedValue}'.");
            }

            return PassSpecialCondition($"Thuá»™c tÃ­nh '{config.PropertyName}' Ä‘Ã£ cÃ³ giÃ¡ trá»‹ Ä‘Ãºng.");
        }

        private static SpecialConditionEvalOutcome EvaluateExcelPrintArea(
            ExcelPrintAreaConfig? config,
            OfficePackage package)
        {
            if (config == null || string.IsNullOrWhiteSpace(config.WorksheetName) || string.IsNullOrWhiteSpace(config.ExpectedRange))
            {
                return FailSpecialCondition("ChÆ°a cáº¥u hÃ¬nh worksheet hoáº·c vÃ¹ng in cáº§n kiá»ƒm tra.");
            }

            if (!package.TryGetXmlDocument("xl/workbook.xml", out var workbookDocument, out var workbookError))
            {
                return FailSpecialCondition(workbookError ?? "KhÃ´ng tÃ¬m tháº¥y xl/workbook.xml trong file há»c sinh.");
            }

            var localSheetId = GetExcelWorksheetLocalSheetId(package, config.WorksheetName);
            if (!localSheetId.HasValue)
            {
                return FailSpecialCondition($"KhÃ´ng tÃ¬m tháº¥y worksheet '{config.WorksheetName}'.");
            }

            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var printAreas = workbookDocument
                .Descendants(x + "definedName")
                .Where(item => string.Equals(item.Attribute("name")?.Value, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase))
                .Where(item => int.TryParse(item.Attribute("localSheetId")?.Value, out var sheetId) && sheetId == localSheetId.Value)
                .Select(item => NormalizeExcelFormulaReference(item.Value))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList();

            if (printAreas.Count == 0)
            {
                return FailSpecialCondition($"Worksheet '{config.WorksheetName}' chÆ°a Ä‘Æ°á»£c Ä‘áº·t Print Area.");
            }

            var expected = NormalizeExcelFormulaReference(config.ExpectedRange);
            var hasMatch = printAreas.Any(area => ExcelDefinedNameRangeMatches(area, expected));
            if (!hasMatch)
            {
                return FailSpecialCondition($"Print Area cá»§a worksheet '{config.WorksheetName}' chÆ°a Ä‘Ãºng vÃ¹ng {config.ExpectedRange}.");
            }

            if (config.RequireExactRange == true && printAreas.Any(area => !ExcelDefinedNameRangeMatches(area, expected)))
            {
                return FailSpecialCondition($"Print Area cá»§a worksheet '{config.WorksheetName}' cÃ³ thÃªm vÃ¹ng ngoÃ i {config.ExpectedRange}.");
            }

            return PassSpecialCondition($"Worksheet '{config.WorksheetName}' Ä‘Ã£ Ä‘áº·t Ä‘Ãºng Print Area.");
        }

        private static SpecialConditionEvalOutcome PassSpecialCondition(string message)
        {
            return new SpecialConditionEvalOutcome { IsPassed = true, Message = message };
        }

        private static SpecialConditionEvalOutcome FailSpecialCondition(string message)
        {
            return new SpecialConditionEvalOutcome { IsPassed = false, Message = message };
        }

        private static List<string> SplitExcelDefinedNameRanges(string formula)
        {
            return formula
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(NormalizeExcelFormulaReference)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool ExcelDefinedNameRangeMatches(string actualRange, string expectedRange)
        {
            return string.Equals(actualRange, expectedRange, StringComparison.OrdinalIgnoreCase)
                || string.Equals(GetExcelReferenceAddressPart(actualRange), GetExcelReferenceAddressPart(expectedRange), StringComparison.OrdinalIgnoreCase);
        }

        private static string GetExcelReferenceAddressPart(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference))
            {
                return string.Empty;
            }

            var normalized = NormalizeExcelFormulaReference(reference);
            var bangIndex = normalized.LastIndexOf('!');
            return bangIndex >= 0 ? normalized[(bangIndex + 1)..] : normalized;
        }

        private static string NormalizeExcelFormulaText(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return string.Empty;
            }

            var normalized = formula.Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal))
            {
                normalized = normalized[1..];
            }

            normalized = normalized.Replace("$", string.Empty);
            normalized = Regex.Replace(normalized, @"\s+", string.Empty, RegexOptions.CultureInvariant);
            return normalized.ToUpperInvariant();
        }

        private static bool ExcelFormulaContainsName(string normalizedFormula, string referenceName)
        {
            var normalizedName = referenceName.Trim().ToUpperInvariant();
            return Regex.IsMatch(
                normalizedFormula,
                $@"(?<![A-Z0-9_.]){Regex.Escape(normalizedName)}(?![A-Z0-9_.])",
                RegexOptions.CultureInvariant);
        }

        private static bool ExcelFormulaContainsFunction(string normalizedFormula, string functionName)
        {
            var normalizedName = functionName.Trim().TrimEnd('(').ToUpperInvariant();
            return Regex.IsMatch(
                normalizedFormula,
                $@"(?<![A-Z0-9_.]){Regex.Escape(normalizedName)}\(",
                RegexOptions.CultureInvariant);
        }

        private static string NormalizeExcelFormulaFragment(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return string.Empty;
            }

            var normalized = formula.Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal))
            {
                normalized = normalized[1..];
            }

            return normalized.ToUpperInvariant();
        }

        private static bool ExcelFormulaContainsCellReference(string normalizedFormula)
        {
            return Regex.IsMatch(
                normalizedFormula,
                @"(?<![A-Z0-9_])(?:'[^']+'!|[A-Z][A-Z0-9_ ]*!|\[.+?\])?[A-Z]{1,3}[1-9][0-9]*(?::[A-Z]{1,3}[1-9][0-9]*)?(?![A-Z0-9_])",
                RegexOptions.CultureInvariant);
        }

        private static string GetExcelCellDisplayText(XElement cell, IReadOnlyList<string> sharedStrings)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var type = cell.Attribute("t")?.Value;
            if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase))
            {
                var sharedStringIndexText = cell.Element(x + "v")?.Value;
                return int.TryParse(sharedStringIndexText, out var sharedStringIndex)
                    && sharedStringIndex >= 0
                    && sharedStringIndex < sharedStrings.Count
                        ? sharedStrings[sharedStringIndex]
                        : string.Empty;
            }

            if (string.Equals(type, "inlineStr", StringComparison.OrdinalIgnoreCase)
                || string.Equals(type, "str", StringComparison.OrdinalIgnoreCase))
            {
                var inlineText = string.Concat(cell.Descendants(x + "t").Select(text => text.Value));
                return string.IsNullOrWhiteSpace(inlineText) ? cell.Element(x + "v")?.Value ?? string.Empty : inlineText;
            }

            return cell.Element(x + "v")?.Value ?? string.Empty;
        }

        private static Dictionary<int, int> GetExcelStyleTextRotations(XDocument stylesDocument)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var cellXfs = stylesDocument
                .Root?
                .Element(x + "cellXfs")?
                .Elements(x + "xf")
                .ToList() ?? new List<XElement>();

            var result = new Dictionary<int, int>();
            for (var index = 0; index < cellXfs.Count; index++)
            {
                var alignment = cellXfs[index].Element(x + "alignment");
                if (int.TryParse(alignment?.Attribute("textRotation")?.Value, out var textRotation))
                {
                    result[index] = textRotation;
                }
            }

            return result;
        }

        private sealed class ExcelSortRow
        {
            public int RowIndex { get; init; }
            public Dictionary<int, string> Values { get; init; } = new();
        }

        private sealed class ExcelResolvedSortKey
        {
            public int Column { get; init; }
            public bool Descending { get; init; }
        }

        private static bool TryResolveExcelSortColumns(
            ExcelMultiColumnSortConfig config,
            OfficePackage package,
            XDocument worksheetDocument,
            out List<ExcelResolvedSortKey> sortKeys,
            out string error)
        {
            sortKeys = new List<ExcelResolvedSortKey>();
            error = string.Empty;
            var headerValues = GetExcelRowValues(package, worksheetDocument, config.HeaderRow ?? 1);

            foreach (var key in config.KeyColumns)
            {
                var column = 0;
                if (!string.IsNullOrWhiteSpace(key.Column))
                {
                    column = GetExcelColumnNumber(key.Column);
                }
                else if (!string.IsNullOrWhiteSpace(key.HeaderName))
                {
                    column = headerValues
                        .FirstOrDefault(pair => string.Equals(pair.Value, key.HeaderName, StringComparison.OrdinalIgnoreCase))
                        .Key;
                }

                if (column <= 0)
                {
                    error = $"KhÃ´ng xÃ¡c Ä‘á»‹nh Ä‘Æ°á»£c cá»™t sáº¯p xáº¿p '{key.HeaderName ?? key.Column}'.";
                    return false;
                }

                sortKeys.Add(new ExcelResolvedSortKey
                {
                    Column = column,
                    Descending = key.Descending == true
                });
            }

            return true;
        }

        private static Dictionary<int, string> GetExcelRowValues(
            OfficePackage package,
            XDocument worksheetDocument,
            int rowIndex)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sharedStrings = GetExcelSharedStrings(package);
            var row = worksheetDocument
                .Descendants(x + "row")
                .FirstOrDefault(item => int.TryParse(item.Attribute("r")?.Value, out var currentRow) && currentRow == rowIndex);

            return row?
                .Elements(x + "c")
                .Select(cell => new
                {
                    Column = GetExcelColumnNumber(Regex.Match(cell.Attribute("r")?.Value ?? string.Empty, "^([A-Z]{1,3})", RegexOptions.IgnoreCase).Groups[1].Value),
                    Text = NormalizePlainText(GetExcelCellDisplayText(cell, sharedStrings))
                })
                .Where(item => item.Column > 0)
                .ToDictionary(item => item.Column, item => item.Text)
                ?? new Dictionary<int, string>();
        }

        private static List<ExcelSortRow> GetExcelDataRowsForSort(
            ExcelMultiColumnSortConfig config,
            OfficePackage package,
            XDocument worksheetDocument)
        {
            XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sharedStrings = GetExcelSharedStrings(package);
            var headerRow = config.HeaderRow ?? 1;
            var hasDataRange = TryParseExcelRange(config.DataRange, out var startColumn, out var startRow, out var endColumn, out var endRow);

            return worksheetDocument
                .Descendants(x + "row")
                .Select(row => new
                {
                    Element = row,
                    RowIndex = int.TryParse(row.Attribute("r")?.Value, out var rowIndex) ? rowIndex : 0
                })
                .Where(row => row.RowIndex > headerRow)
                .Where(row => !hasDataRange || row.RowIndex >= startRow && row.RowIndex <= endRow)
                .Select(row => new ExcelSortRow
                {
                    RowIndex = row.RowIndex,
                    Values = row.Element
                        .Elements(x + "c")
                        .Select(cell => new
                        {
                            Column = GetExcelColumnNumber(Regex.Match(cell.Attribute("r")?.Value ?? string.Empty, "^([A-Z]{1,3})", RegexOptions.IgnoreCase).Groups[1].Value),
                            Text = NormalizePlainText(GetExcelCellDisplayText(cell, sharedStrings))
                        })
                        .Where(item => item.Column > 0)
                        .Where(item => !hasDataRange || item.Column >= startColumn && item.Column <= endColumn)
                        .ToDictionary(item => item.Column, item => item.Text)
                })
                .Where(row => row.Values.Values.Any(value => !string.IsNullOrWhiteSpace(value)))
                .ToList();
        }

        private static int CompareExcelSortRows(
            ExcelSortRow left,
            ExcelSortRow right,
            IReadOnlyList<ExcelResolvedSortKey> sortKeys)
        {
            foreach (var sortKey in sortKeys)
            {
                left.Values.TryGetValue(sortKey.Column, out var leftValue);
                right.Values.TryGetValue(sortKey.Column, out var rightValue);
                var compare = CompareExcelSortValues(leftValue ?? string.Empty, rightValue ?? string.Empty);
                if (compare != 0)
                {
                    return sortKey.Descending ? -compare : compare;
                }
            }

            return 0;
        }

        private static int CompareExcelSortValues(string left, string right)
        {
            if (decimal.TryParse(left, NumberStyles.Number, CultureInfo.InvariantCulture, out var leftNumber)
                && decimal.TryParse(right, NumberStyles.Number, CultureInfo.InvariantCulture, out var rightNumber))
            {
                return leftNumber.CompareTo(rightNumber);
            }

            return string.Compare(left, right, StringComparison.CurrentCultureIgnoreCase);
        }

        private static bool ExcelDecimalAttributeEquals(string? actualValue, decimal expectedValue)
        {
            var actual = ParseExcelDecimal(actualValue) ?? 0m;
            return actual == expectedValue;
        }

        private static decimal? ParseExcelDecimal(string? value)
        {
            return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var result)
                ? result
                : null;
        }

        private static int? GetExcelWorksheetLocalSheetId(OfficePackage package, string worksheetName)
        {
            var worksheets = GetExcelWorksheets(package);
            for (var index = 0; index < worksheets.Count; index++)
            {
                if (string.Equals(worksheets[index].Name, worksheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return index;
                }
            }

            return null;
        }

        private static XDocument? TryParsePackageXml(OfficePackage package, string sourceFile)
        {
            return package.TryGetXmlDocument(sourceFile, out var document, out _)
                ? document
                : null;
        }

        private static string GetRelationshipPartPath(string sourceFile)
        {
            var normalized = NormalizeSourceFile(sourceFile);
            var slashIndex = normalized.LastIndexOf('/');
            if (slashIndex < 0)
            {
                return $"_rels/{normalized}.rels";
            }

            var folder = normalized[..slashIndex];
            var fileName = normalized[(slashIndex + 1)..];
            return $"{folder}/_rels/{fileName}.rels";
        }

        private static string GetPackageFolder(string sourceFile)
        {
            var normalized = NormalizeSourceFile(sourceFile);
            var slashIndex = normalized.LastIndexOf('/');
            return slashIndex < 0 ? string.Empty : normalized[..slashIndex];
        }

        private static string NormalizeRelationshipTarget(string baseFolder, string target)
        {
            if (string.IsNullOrWhiteSpace(target))
            {
                return string.Empty;
            }

            var normalizedTarget = target.Trim().Replace('\\', '/');
            if (normalizedTarget.StartsWith("/", StringComparison.Ordinal))
            {
                return NormalizeSourceFile(normalizedTarget);
            }

            var combined = string.IsNullOrWhiteSpace(baseFolder)
                ? normalizedTarget
                : $"{baseFolder.TrimEnd('/')}/{normalizedTarget}";

            var segments = new List<string>();
            foreach (var segment in combined.Split('/', StringSplitOptions.RemoveEmptyEntries))
            {
                if (segment == ".")
                {
                    continue;
                }

                if (segment == "..")
                {
                    if (segments.Count > 0)
                    {
                        segments.RemoveAt(segments.Count - 1);
                    }

                    continue;
                }

                segments.Add(segment);
            }

            return string.Join("/", segments);
        }

        private static bool TryParseExcelRange(
            string? range,
            out int startColumn,
            out int startRow,
            out int endColumn,
            out int endRow)
        {
            startColumn = 0;
            startRow = 0;
            endColumn = 0;
            endRow = 0;

            if (string.IsNullOrWhiteSpace(range))
            {
                return false;
            }

            var match = Regex.Match(
                range.Trim(),
                @"^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return false;
            }

            startColumn = GetExcelColumnNumber(match.Groups[1].Value);
            startRow = int.Parse(match.Groups[2].Value);
            endColumn = match.Groups[3].Success ? GetExcelColumnNumber(match.Groups[3].Value) : startColumn;
            endRow = match.Groups[4].Success ? int.Parse(match.Groups[4].Value) : startRow;

            if (startColumn <= 0 || endColumn <= 0 || startRow <= 0 || endRow <= 0)
            {
                return false;
            }

            if (startColumn > endColumn)
            {
                (startColumn, endColumn) = (endColumn, startColumn);
            }

            if (startRow > endRow)
            {
                (startRow, endRow) = (endRow, startRow);
            }

            return true;
        }

        private static int GetExcelColumnNumber(string columnName)
        {
            var result = 0;
            foreach (var ch in columnName.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z')
                {
                    return 0;
                }

                result = (result * 26) + (ch - 'A' + 1);
            }

            return result;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            var name = new StringBuilder();
            var number = columnNumber;
            while (number > 0)
            {
                number--;
                name.Insert(0, (char)('A' + (number % 26)));
                number /= 26;
            }

            return name.ToString();
        }

        private static SpecialConditionEvalOutcome EvaluateTextBoxContainsText(
            TextBoxContainsTextConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh TextBox Contains Text (textBoxContainsTextConfig trong).");
            }

            var expectedText = NormalizePlainText(config.ExpectedText);
            if (string.IsNullOrWhiteSpace(expectedText))
            {
                return Fail("textBoxContainsTextConfig.expectedText khong duoc rong.");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);

            if (!package.TryGetXmlDocument(sourceFile, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            try
            {
                var comparison = config.CaseSensitive == true
                    ? StringComparison.Ordinal
                    : StringComparison.OrdinalIgnoreCase;
                var matchMode = string.IsNullOrWhiteSpace(config.MatchMode) ? "exact" : config.MatchMode.Trim();
                var occurrence = Math.Max(1, config.TargetOccurrence ?? 1);

                var textBoxes = document.Descendants(w + "txbxContent").ToList();
                if (textBoxes.Count == 0)
                {
                    return Fail("Khong tim thay textbox nao trong tai lieu.");
                }

                var matches = textBoxes
                    .Select((textBox, index) => new
                    {
                        TextBox = textBox,
                        Index = index + 1,
                        Text = NormalizePlainText(string.Join(
                            " ",
                            textBox.Descendants(w + "p")
                                .Select(paragraph => BuildParagraphTextSnapshot(paragraph, w).Text)
                                .Where(text => !string.IsNullOrWhiteSpace(text))))
                    })
                    .Where(item => TextMatches(item.Text, expectedText, matchMode, comparison))
                    .ToList();

                if (matches.Count < occurrence)
                {
                    return Fail($"Chi tim thay {matches.Count} textbox chua dung noi dung can cham, can occurrence {occurrence}.");
                }

                var target = matches[occurrence - 1];

                if (config.RequireDefaultPaste != false && HasNonDefaultTextBoxParagraphStyle(target.TextBox, w))
                {
                    return Fail($"Textbox thu {target.Index} co paragraph style rieng, co the khong phai paste mac dinh.");
                }

                if (config.RequireDefaultPaste != false
                    && TryGetForbiddenTextBoxRunProperty(target.TextBox, w, config.ForbiddenRunProperties, out var forbiddenProperty))
                {
                    return Fail($"Textbox thu {target.Index} co dinh dang '{forbiddenProperty}', co the da Paste Merge Formatting.");
                }

                if (config.RequireDefaultPaste != false
                    && TryGetForbiddenTextBoxColor(target.TextBox, w, config.ForbiddenTextColors, out var forbiddenColor))
                {
                    return Fail($"Textbox thu {target.Index} co mau chu '{forbiddenColor}' giong style cua textbox, co the da Paste Merge Formatting.");
                }

                if (config.RequireRemovedFromBody != false)
                {
                    var bodyParagraphsOutsideTextBoxes = document
                        .Descendants(w + "body")
                        .Elements(w + "p")
                        .Where(paragraph => !paragraph.Ancestors(w + "txbxContent").Any())
                        .Select(paragraph => BuildParagraphTextSnapshot(paragraph, w, excludeTextBoxContent: true).Text)
                        .Where(text => !string.IsNullOrWhiteSpace(text))
                        .ToList();

                    if (bodyParagraphsOutsideTextBoxes.Any(text => TextMatches(text, expectedText, matchMode, comparison)))
                    {
                        return Fail("Doan van can dua vao textbox van con nam ngoai than tai lieu, co the hoc sinh copy thay vi cut.");
                    }
                }

                return new SpecialConditionEvalOutcome
                {
                    IsPassed = true,
                    Message = $"Textbox thu {target.Index} chua dung noi dung yeu cau."
                };
            }
            catch (XmlException ex)
            {
                return Fail($"Khong the phan tich XML: {ex.Message}");
            }
        }

        private static bool TextMatches(string actualText, string expectedText, string matchMode, StringComparison comparison)
        {
            var actual = NormalizePlainText(actualText);
            var expected = NormalizePlainText(expectedText);

            return string.Equals(matchMode, "contains", StringComparison.OrdinalIgnoreCase)
                ? actual.Contains(expected, comparison)
                : string.Equals(actual, expected, comparison);
        }

        private static bool HasNonDefaultTextBoxParagraphStyle(XElement textBoxContent, XNamespace w)
        {
            foreach (var paragraph in textBoxContent.Descendants(w + "p"))
            {
                var style = paragraph.Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val")?.Value;
                if (!string.IsNullOrWhiteSpace(style)
                    && !string.Equals(style, "Normal", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetForbiddenTextBoxRunProperty(
            XElement textBoxContent,
            XNamespace w,
            IReadOnlyList<string>? configuredForbiddenProperties,
            out string? forbiddenProperty)
        {
            forbiddenProperty = null;
            var defaultTextBoxProperties = new HashSet<string>(
                new[] { "caps", "szCs", "noProof" },
                StringComparer.OrdinalIgnoreCase);
            var forbiddenProperties = new HashSet<string>(
                (configuredForbiddenProperties ?? Array.Empty<string>())
                    .Where(property => !defaultTextBoxProperties.Contains(property)),
                StringComparer.OrdinalIgnoreCase);

            if (forbiddenProperties.Count == 0)
            {
                return false;
            }

            var runPropertyContainers = textBoxContent
                .Descendants(w + "p")
                .Select(paragraph => paragraph.Element(w + "pPr")?.Element(w + "rPr"))
                .Concat(textBoxContent.Descendants(w + "r").Select(run => run.Element(w + "rPr")))
                .Where(runProperties => runProperties != null)
                .Cast<XElement>();

            foreach (var runProperties in runPropertyContainers)
            {
                var matched = runProperties.Elements()
                    .Select(element => element.Name.LocalName)
                    .FirstOrDefault(name => forbiddenProperties.Contains(name));

                if (!string.IsNullOrWhiteSpace(matched))
                {
                    forbiddenProperty = matched;
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetForbiddenTextBoxColor(
            XElement textBoxContent,
            XNamespace w,
            IReadOnlyList<string>? configuredForbiddenColors,
            out string? forbiddenColor)
        {
            forbiddenColor = null;
            var forbiddenColors = new HashSet<string>(
                (configuredForbiddenColors == null || configuredForbiddenColors.Count == 0
                    ? new[] { "FFFFFF", "background1", "bg1", "lt1" }
                    : configuredForbiddenColors)
                .Select(NormalizeHexColor)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!),
                StringComparer.OrdinalIgnoreCase);

            foreach (var run in textBoxContent.Descendants(w + "r"))
            {
                var runText = NormalizePlainText(string.Concat(run.Descendants(w + "t").Select(text => text.Value)));
                if (string.IsNullOrWhiteSpace(runText))
                {
                    continue;
                }

                var color = run.Element(w + "rPr")?.Element(w + "color");
                var colorValues = new[]
                {
                    color?.Attribute(w + "val")?.Value,
                    color?.Attribute(w + "themeColor")?.Value
                }
                    .Select(NormalizeHexColor)
                    .Where(value => !string.IsNullOrWhiteSpace(value));

                var matchedColor = colorValues.FirstOrDefault(value => forbiddenColors.Contains(value!));
                if (!string.IsNullOrWhiteSpace(matchedColor))
                {
                    forbiddenColor = matchedColor;
                    return true;
                }
            }

            return false;
        }

        private static SpecialConditionEvalOutcome EvaluatePageMargins(
            PageMarginsConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Page Margins (pageMarginsConfig trong).");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);

            if (!package.TryGetXmlDocument(sourceFile, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            try
            {
                var margins = document.Descendants(w + "sectPr")
                    .Select(section => section.Element(w + "pgMar"))
                    .Where(pgMar => pgMar != null)
                    .Cast<XElement>()
                    .ToList();

                if (margins.Count == 0)
                {
                    return Fail("Khong tim thay w:pgMar trong section properties.");
                }

                var requireAllSections = config.RequireAllSections != false;
                var checkedMargins = requireAllSections ? margins : margins.TakeLast(1).ToList();
                var sectionNumber = requireAllSections ? 1 : margins.Count;

                foreach (var margin in checkedMargins)
                {
                    var mismatch = GetMarginMismatch(margin, w, config);
                    if (mismatch != null)
                    {
                        return Fail(requireAllSections
                            ? $"Section {sectionNumber}: {mismatch}"
                            : mismatch);
                    }

                    sectionNumber++;
                }

                return new SpecialConditionEvalOutcome
                {
                    IsPassed = true,
                    Message = requireAllSections
                        ? $"Tat ca {margins.Count} section co margin dung yeu cau."
                        : "Section cuoi co margin dung yeu cau."
                };
            }
            catch (XmlException ex)
            {
                return Fail($"Khong the phan tich XML: {ex.Message}");
            }
        }

        private static string? GetMarginMismatch(XElement margin, XNamespace w, PageMarginsConfig config)
        {
            string? Check(string attributeName, int? expected)
            {
                if (!expected.HasValue)
                {
                    return null;
                }

                var actualValue = margin.Attribute(w + attributeName)?.Value;
                if (!int.TryParse(actualValue, out var actual))
                {
                    return $"Khong doc duoc margin {attributeName}.";
                }

                return actual == expected.Value
                    ? null
                    : $"Margin {attributeName} la {actual}, can {expected.Value}.";
            }

            return Check("top", config.Top)
                ?? Check("bottom", config.Bottom)
                ?? Check("left", config.Left)
                ?? Check("right", config.Right)
                ?? Check("gutter", config.Gutter);
        }

        private static SpecialConditionEvalOutcome EvaluatePageBorder(
            PageBorderConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Page Border (pageBorderConfig trong).");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);

            if (!package.TryGetXmlDocument(sourceFile, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            try
            {
                var pageBorders = document.Descendants(w + "sectPr")
                    .Select(section => section.Element(w + "pgBorders"))
                    .Where(pgBorders => pgBorders != null)
                    .Cast<XElement>()
                    .ToList();

                if (pageBorders.Count == 0)
                {
                    return Fail("Khong tim thay w:pgBorders trong section properties.");
                }

                var requireAllSections = config.RequireAllSections != false;
                var checkedBorders = requireAllSections ? pageBorders : pageBorders.TakeLast(1).ToList();
                var sectionNumber = requireAllSections ? 1 : pageBorders.Count;

                foreach (var pageBorder in checkedBorders)
                {
                    var mismatch = GetPageBorderMismatch(pageBorder, w, config);
                    if (mismatch != null)
                    {
                        return Fail(requireAllSections
                            ? $"Section {sectionNumber}: {mismatch}"
                            : mismatch);
                    }

                    sectionNumber++;
                }

                return new SpecialConditionEvalOutcome
                {
                    IsPassed = true,
                    Message = requireAllSections
                        ? $"Tat ca {pageBorders.Count} section co page border dung yeu cau."
                        : "Section cuoi co page border dung yeu cau."
                };
            }
            catch (XmlException ex)
            {
                return Fail($"Khong the phan tich XML: {ex.Message}");
            }
        }

        private static string? GetPageBorderMismatch(XElement pageBorder, XNamespace w, PageBorderConfig config)
        {
            var requiredSides = config.RequireBox != false
                ? new[] { "top", "left", "bottom", "right" }
                : pageBorder.Elements().Select(element => element.Name.LocalName).ToArray();

            if (requiredSides.Length == 0)
            {
                return "Khong co canh border nao de kiem tra.";
            }

            foreach (var sideName in requiredSides)
            {
                var side = pageBorder.Element(w + sideName);
                if (side == null)
                {
                    return $"Thieu border canh {sideName}.";
                }

                var styleMismatch = GetPageBorderStyleMismatch(side, w, config.RequiredStyle);
                if (styleMismatch != null)
                {
                    return $"Canh {sideName}: {styleMismatch}";
                }

                var widthMismatch = GetPageBorderWidthMismatch(side, w, config.RequiredWidth, config.MinWidth);
                if (widthMismatch != null)
                {
                    return $"Canh {sideName}: {widthMismatch}";
                }

                var colorMismatch = GetPageBorderColorMismatch(side, w, config);
                if (colorMismatch != null)
                {
                    return $"Canh {sideName}: {colorMismatch}";
                }
            }

            return null;
        }

        private static string? GetPageBorderStyleMismatch(XElement borderSide, XNamespace w, string? expectedStyle)
        {
            if (string.IsNullOrWhiteSpace(expectedStyle))
            {
                return null;
            }

            var actualStyle = borderSide.Attribute(w + "val")?.Value;
            if (string.Equals(actualStyle, expectedStyle.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return $"style la '{actualStyle ?? "(rong)"}', can '{expectedStyle.Trim()}'.";
        }

        private static string? GetPageBorderWidthMismatch(XElement borderSide, XNamespace w, int? requiredWidth, int? minWidth)
        {
            if (!requiredWidth.HasValue && !minWidth.HasValue)
            {
                return null;
            }

            var actualWidthText = borderSide.Attribute(w + "sz")?.Value;
            if (!int.TryParse(actualWidthText, out var actualWidth))
            {
                return "khong doc duoc do rong border w:sz.";
            }

            if (requiredWidth.HasValue && actualWidth != requiredWidth.Value)
            {
                return $"do rong la {actualWidth}, can {requiredWidth.Value}.";
            }

            if (minWidth.HasValue && actualWidth < minWidth.Value)
            {
                return $"do rong la {actualWidth}, can toi thieu {minWidth.Value}.";
            }

            return null;
        }

        private static string? GetPageBorderColorMismatch(XElement borderSide, XNamespace w, PageBorderConfig config)
        {
            var expectedColors = config.AllowedColors?
                .Select(NormalizeBorderColor)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .ToList() ?? new List<string>();

            if (!string.IsNullOrWhiteSpace(config.RequiredColor))
            {
                expectedColors.Insert(0, NormalizeBorderColor(config.RequiredColor)!);
            }

            expectedColors = expectedColors
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (expectedColors.Count == 0)
            {
                return null;
            }

            var actualColors = new[]
            {
                borderSide.Attribute(w + "color")?.Value,
                borderSide.Attribute(w + "themeColor")?.Value
            }
                .Select(NormalizeBorderColor)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .ToList();

            if (actualColors.Count == 0)
            {
                return "khong doc duoc mau border.";
            }

            if (actualColors.Any(actual => expectedColors.Any(expected => DoesBorderColorMatch(actual, expected))))
            {
                return null;
            }

            return $"mau la '{string.Join("/", actualColors)}', can mot trong '{string.Join("/", expectedColors)}'.";
        }

        private static SpecialConditionEvalOutcome EvaluateDocumentStyleSet(
            DocumentStyleSetConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Document Style Set (documentStyleSetConfig trong).");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/styles.xml"
                : NormalizeSourceFile(config.SourceFile);

            if (!package.XmlParts.TryGetValue(sourceFile, out var actualXml))
            {
                return Fail($"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            var expectedFragments = config.ExpectedFragments?
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList() ?? new List<string>();

            if (expectedFragments.Count == 0)
            {
                return Fail("documentStyleSetConfig.expectedFragments khong duoc rong.");
            }

            var matchPolicy = string.IsNullOrWhiteSpace(config.MatchPolicy)
                ? XmlGradingMatchPolicies.All
                : config.MatchPolicy.Trim();
            var cache = new XmlEvaluationCache();
            var matches = expectedFragments
                .Select(fragment => MatchStyleSetFragment(sourceFile, actualXml, fragment, config.IgnoreAttributes, cache))
                .ToList();
            var passed = string.Equals(matchPolicy, XmlGradingMatchPolicies.Any, StringComparison.OrdinalIgnoreCase)
                ? matches.Any(match => match.IsMatched)
                : matches.All(match => match.IsMatched);

            if (!passed)
            {
                var missing = matches.Count(match => !match.IsMatched);
                var styleName = string.IsNullOrWhiteSpace(config.StyleSetName) ? "style set" : config.StyleSetName!.Trim();
                return Fail($"{styleName}: thieu {missing}/{matches.Count} dau hieu XML trong {sourceFile}.");
            }

            return new SpecialConditionEvalOutcome
            {
                IsPassed = true,
                Message = string.IsNullOrWhiteSpace(config.StyleSetName)
                    ? $"Document style set khop {matches.Count} dau hieu XML."
                    : $"Document style set '{config.StyleSetName}' khop {matches.Count} dau hieu XML."
            };
        }

        private static ExpectedMatchResult MatchStyleSetFragment(
            string sourceFile,
            string actualXml,
            string expectedFragment,
            IReadOnlyList<string>? ignoreAttributes,
            XmlEvaluationCache cache)
        {
            var trimmed = expectedFragment.Trim();
            if (trimmed.StartsWith("<", StringComparison.Ordinal))
            {
                return XmlContainsNormalized(sourceFile, actualXml, trimmed, ignoreAttributes ?? Array.Empty<string>(), cache);
            }

            return RawContains(actualXml, trimmed, trim: true);
        }

        private static SpecialConditionEvalOutcome EvaluateSectionBreakBeforeText(
            SectionBreakBeforeTextConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Section Break Before Text (sectionBreakBeforeTextConfig trong).");
            }

            var targetText = NormalizePlainText(config.TargetText);
            if (string.IsNullOrWhiteSpace(targetText))
            {
                return Fail("sectionBreakBeforeTextConfig.targetText khong duoc rong.");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);

            if (!package.TryGetXmlDocument(sourceFile, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            try
            {
                var paragraphs = document.Descendants(w + "body").Elements(w + "p").ToList();
                var matches = paragraphs
                    .Select((paragraph, index) => new
                    {
                        Paragraph = paragraph,
                        Index = index,
                        Text = BuildParagraphTextSnapshot(paragraph, w).Text
                    })
                    .Where(item => item.Text.Contains(targetText, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (matches.Count == 0)
                {
                    return Fail($"Khong tim thay targetText '{targetText}' trong {sourceFile}.");
                }

                var occurrence = Math.Max(1, config.TargetOccurrence ?? 1);
                if (matches.Count < occurrence)
                {
                    return Fail($"Chi tim thay {matches.Count} paragraph chua targetText '{targetText}', can occurrence {occurrence}.");
                }

                var target = matches[occurrence - 1];
                var candidateParagraphs = new List<(XElement Paragraph, string Position)>();

                if (config.AllowSameParagraphSectPr != false)
                {
                    candidateParagraphs.Add((target.Paragraph, "cung paragraph voi targetText"));
                }

                if (target.Index > 0)
                {
                    if (config.RequireImmediateBefore != false)
                    {
                        candidateParagraphs.Add((paragraphs[target.Index - 1], "paragraph ngay truoc targetText"));
                    }
                    else
                    {
                        candidateParagraphs.AddRange(
                            paragraphs.Take(target.Index).Reverse().Select(paragraph => (paragraph, "paragraph nam truoc targetText")));
                    }
                }

                var expectedType = string.IsNullOrWhiteSpace(config.BreakType)
                    ? "continuous"
                    : config.BreakType.Trim();

                foreach (var candidate in candidateParagraphs)
                {
                    var sectPr = candidate.Paragraph.Element(w + "pPr")?.Element(w + "sectPr");
                    if (sectPr == null)
                    {
                        continue;
                    }

                    var actualType = sectPr.Element(w + "type")?.Attribute(w + "val")?.Value;
                    actualType = string.IsNullOrWhiteSpace(actualType) ? "nextPage" : actualType.Trim();

                    if (string.Equals(actualType, expectedType, StringComparison.OrdinalIgnoreCase))
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = $"Da tim thay section break '{expectedType}' {candidate.Position}."
                        };
                    }

                    return Fail($"Tim thay section break {candidate.Position} nhung type la '{actualType}' thay vi '{expectedType}'.");
                }

                return Fail($"Khong tim thay section break '{expectedType}' ngay truoc targetText '{targetText}'.");
            }
            catch (XmlException ex)
            {
                return Fail($"Khong the phan tich XML: {ex.Message}");
            }
        }

        private static SpecialConditionEvalOutcome EvaluatePictureStyle(
            PictureStyleConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Picture Style (pictureStyleConfig trong).");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);

            var relsFile = string.IsNullOrWhiteSpace(config.RelsFile)
                ? "word/_rels/document.xml.rels"
                : NormalizeSourceFile(config.RelsFile);

            if (!package.TryGetXmlDocument(sourceFile, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";

            try
            {
                var drawings = document.Descendants(w + "drawing").ToList();
                if (drawings.Count == 0)
                {
                    return Fail("Khong tim thay hinh anh (w:drawing) nao trong tai lieu.");
                }

                var expectedHash = string.IsNullOrWhiteSpace(config.ImageHash)
                    ? null
                    : ImageHashUtility.NormalizeHash(config.ImageHash);
                var expectedPerceptualHash = config.PerceptualHash;
                var targetImageIndex = Math.Max(1, config.TargetImageIndex ?? 1);
                var stylePreset = string.IsNullOrWhiteSpace(config.StylePreset)
                    ? "simpleFrameBlack"
                    : config.StylePreset.Trim();
                var expectedLineColor = NormalizeHexColor(config.RequiredLineColor);
                var expectedGeometry = string.IsNullOrWhiteSpace(config.PresetGeometry)
                    ? null
                    : config.PresetGeometry.Trim();

                var imageOrdinal = 0;
                string? lastMismatchInfo = null;
                Dictionary<string, string>? relationships = null;

                if (!string.IsNullOrWhiteSpace(expectedHash)
                    && !package.TryGetRelationships(relsFile, out relationships, out var relsError))
                {
                    return Fail(relsError ?? $"Khong tim thay {relsFile} trong file hoc sinh.");
                }

                foreach (var drawing in drawings)
                {
                    imageOrdinal++;

                    if (!string.IsNullOrWhiteSpace(expectedHash))
                    {
                        var relationshipId = drawing.Descendants(a + "blip").FirstOrDefault()?.Attribute(r + "embed")?.Value;
                        if (string.IsNullOrWhiteSpace(relationshipId))
                        {
                            lastMismatchInfo = $"Anh thu {imageOrdinal} khong co relationship id.";
                            continue;
                        }

                        if (relationships == null
                            || !relationships.TryGetValue(relationshipId, out var target)
                            || string.IsNullOrWhiteSpace(target))
                        {
                            lastMismatchInfo = $"Relationship {relationshipId} khong co Target hop le.";
                            continue;
                        }

                        var imagePath = ResolveRelationshipTarget(sourceFile, target);
                        if (!package.BinaryParts.TryGetValue(imagePath, out var imageBytes))
                        {
                            lastMismatchInfo = $"Khong doc duoc anh {imagePath} trong file hoc sinh.";
                            continue;
                        }

                        if (!IsImageMatch(package, imagePath, imageBytes, expectedHash, expectedPerceptualHash))
                        {
                            lastMismatchInfo = $"Anh thu {imageOrdinal} khong dung noi dung anh muc tieu.";
                            continue;
                        }
                    }
                    else if (imageOrdinal != targetImageIndex)
                    {
                        continue;
                    }

                    var picture = drawing.Descendants(pic + "pic").FirstOrDefault();
                    var shapeProperties = picture?.Element(pic + "spPr");
                    if (shapeProperties == null)
                    {
                        return Fail($"Anh thu {imageOrdinal} khong co pic:spPr de kiem tra picture style.");
                    }

                    var line = shapeProperties.Element(a + "ln");
                    if (line == null)
                    {
                        return Fail($"Anh thu {imageOrdinal} chua co vien anh (a:ln).");
                    }

                    if (!string.IsNullOrWhiteSpace(expectedLineColor))
                    {
                        var actualColor = GetLineColor(line, a);
                        if (!DoesLineColorMatch(actualColor, expectedLineColor))
                        {
                            return Fail($"Vien anh thu {imageOrdinal} co mau '{actualColor ?? "unknown"}' thay vi '{expectedLineColor}'.");
                        }
                    }

                    if (config.MinLineWidth.HasValue)
                    {
                        var actualWidth = ParseIntAttribute(line, "w") ?? 0;
                        if (actualWidth < config.MinLineWidth.Value)
                        {
                            return Fail($"Vien anh thu {imageOrdinal} co do day {actualWidth}, can toi thieu {config.MinLineWidth.Value}.");
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(expectedGeometry))
                    {
                        var actualGeometry = shapeProperties.Element(a + "prstGeom")?.Attribute("prst")?.Value;
                        if (!string.Equals(actualGeometry, expectedGeometry, StringComparison.OrdinalIgnoreCase))
                        {
                            return Fail($"Anh thu {imageOrdinal} co geometry '{actualGeometry ?? "unknown"}' thay vi '{expectedGeometry}'.");
                        }
                    }

                    if (string.Equals(stylePreset, "simpleFrameBlack", StringComparison.OrdinalIgnoreCase)
                        && HasVisiblePictureEffects(picture!, shapeProperties, a, pic))
                    {
                        return Fail($"Anh thu {imageOrdinal} co effect/shadow nen khong phai Simple Frame, Black.");
                    }

                    return new SpecialConditionEvalOutcome
                    {
                        IsPassed = true,
                        Message = $"Anh thu {imageOrdinal} da co picture style phu hop."
                    };
                }

                return Fail(lastMismatchInfo ?? $"Khong tim thay anh muc tieu de kiem tra picture style.");
            }
            catch (XmlException ex)
            {
                return Fail($"Khong the phan tich XML: {ex.Message}");
            }
        }

        private static SpecialConditionEvalOutcome EvaluateHyperlink(
            HyperlinkConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("Chua cau hinh Hyperlink (hyperlinkConfig trong).");
            }

            var expectedText = NormalizePlainText(config.DisplayText);
            var expectedUrl = config.Url?.Trim();

            if (string.IsNullOrWhiteSpace(expectedText))
            {
                return Fail("hyperlinkConfig.displayText khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(expectedUrl))
            {
                return Fail("hyperlinkConfig.url khong duoc rong.");
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);

            var relsFile = string.IsNullOrWhiteSpace(config.RelsFile)
                ? "word/_rels/document.xml.rels"
                : NormalizeSourceFile(config.RelsFile);

            if (!package.TryGetXmlDocument(sourceFile, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            if (!package.TryGetRelationships(relsFile, out var relationships, out var relsError))
            {
                return Fail(relsError ?? $"Khong tim thay {relsFile} trong file hoc sinh.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            try
            {
                var textComparison = config.CaseSensitiveText == true
                    ? StringComparison.Ordinal
                    : StringComparison.OrdinalIgnoreCase;

                var anchorTextBefore = NormalizePlainText(config.AnchorTextBefore);
                string? matchingTextWrongUrl = null;
                var sawAnchorParagraph = string.IsNullOrWhiteSpace(anchorTextBefore);
                var candidateHyperlinks = document.Descendants(w + "p")
                    .Where(paragraph =>
                    {
                        if (string.IsNullOrWhiteSpace(anchorTextBefore))
                        {
                            return true;
                        }

                        var paragraphText = BuildParagraphTextSnapshot(paragraph, w).Text;
                        var anchorIndex = paragraphText.IndexOf(anchorTextBefore, textComparison);
                        if (anchorIndex < 0)
                        {
                            return false;
                        }

                        sawAnchorParagraph = true;
                        var expectedTextIndex = paragraphText.IndexOf(
                            expectedText,
                            anchorIndex + anchorTextBefore.Length,
                            textComparison);

                        return expectedTextIndex >= 0;
                    })
                    .SelectMany(paragraph => paragraph.Descendants(w + "hyperlink"));

                foreach (var hyperlink in candidateHyperlinks)
                {
                    var relationshipId = hyperlink.Attribute(r + "id")?.Value;
                    if (string.IsNullOrWhiteSpace(relationshipId))
                    {
                        continue;
                    }

                    var actualText = NormalizePlainText(string.Concat(hyperlink.Descendants(w + "t").Select(text => text.Value)));
                    if (!string.Equals(actualText, expectedText, textComparison))
                    {
                        continue;
                    }

                    if (!relationships.TryGetValue(relationshipId, out var target) || string.IsNullOrWhiteSpace(target))
                    {
                        matchingTextWrongUrl = $"Tim thay hyperlink text '{actualText}' nhung khong co relationship target.";
                        continue;
                    }

                    var actualUrl = target.Trim();
                    if (string.Equals(actualUrl, expectedUrl, StringComparison.OrdinalIgnoreCase))
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = $"Da tao hyperlink dung cho '{expectedText}'."
                        };
                    }

                    matchingTextWrongUrl = $"Tim thay hyperlink text '{actualText}' nhung URL la '{actualUrl}' thay vi '{expectedUrl}'.";
                }

                if (!sawAnchorParagraph)
                {
                    return Fail($"Khong tim thay paragraph chua anchorTextBefore '{anchorTextBefore}'.");
                }

                if (!string.IsNullOrWhiteSpace(anchorTextBefore) && matchingTextWrongUrl == null)
                {
                    return Fail($"Khong tim thay hyperlink cho text '{expectedText}' sau anchorTextBefore '{anchorTextBefore}'.");
                }

                return Fail(matchingTextWrongUrl ?? $"Khong tim thay hyperlink cho text '{expectedText}'.");
            }
            catch (XmlException ex)
            {
                return Fail($"Khong the phan tich XML: {ex.Message}");
            }
        }

        private static SpecialConditionEvalOutcome EvaluateConvertTableToText(
            ConvertTableToTextConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            config ??= new ConvertTableToTextConfig();

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);

            if (!package.TryGetXmlDocument(sourceFile, out var document, out var documentError))
            {
                return Fail(documentError ?? $"Khong tim thay {sourceFile} trong file hoc sinh.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            try
            {
                if (config.RequireNoTables != false && document.Descendants(w + "tbl").Any())
                {
                    return Fail("Tai lieu van con bang Word (w:tbl), chua convert table thanh text.");
                }

                var paragraphs = document
                    .Descendants(w + "p")
                    .Select(paragraph => BuildParagraphTextSnapshot(paragraph, w))
                    .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                    .ToList();

                var startIndex = 0;
                var anchorText = NormalizePlainText(config.AnchorText);
                if (!string.IsNullOrWhiteSpace(anchorText))
                {
                    var anchorIndex = paragraphs.FindIndex(paragraph =>
                        NormalizePlainText(paragraph.Text).Contains(anchorText, StringComparison.OrdinalIgnoreCase));

                    if (anchorIndex < 0)
                    {
                        return Fail($"Khong tim thay anchorText '{anchorText}'.");
                    }

                    startIndex = anchorIndex + 1;
                }

                var minTabsPerRow = Math.Max(1, config.MinTabsPerRow ?? 1);
                var convertedRows = paragraphs
                    .Skip(startIndex)
                    .Where(paragraph => paragraph.TabCount >= minTabsPerRow)
                    .ToList();

                var expectedRows = config.ExpectedRows?
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(NormalizePlainText)
                    .ToList() ?? new List<string>();

                if (expectedRows.Count > 0)
                {
                    var cursor = 0;
                    foreach (var expectedRow in expectedRows)
                    {
                        var index = convertedRows.FindIndex(cursor, row =>
                            string.Equals(NormalizePlainText(row.Text), expectedRow, StringComparison.OrdinalIgnoreCase));

                        if (index < 0)
                        {
                            return Fail($"Khong tim thay dong da convert: {expectedRow}");
                        }

                        cursor = index + 1;
                    }
                }

                var requiredRows = Math.Max(expectedRows.Count, config.MinRows ?? 1);
                if (convertedRows.Count < requiredRows)
                {
                    return Fail($"Chi tim thay {convertedRows.Count} dong text co tab, can toi thieu {requiredRows} dong.");
                }

                return new SpecialConditionEvalOutcome
                {
                    IsPassed = true,
                    Message = $"Da convert table thanh text bang tabs: tim thay {convertedRows.Count} dong phu hop."
                };
            }
            catch (XmlException ex)
            {
                return Fail($"Khong the phan tich XML: {ex.Message}");
            }
        }

        private static ParagraphTextSnapshot BuildParagraphTextSnapshot(
            XElement paragraph,
            XNamespace w,
            bool excludeTextBoxContent = false)
        {
            var text = new StringBuilder();
            var tabCount = 0;

            foreach (var node in paragraph.Descendants())
            {
                if (excludeTextBoxContent && node.Ancestors(w + "txbxContent").Any())
                {
                    continue;
                }

                if (node.Name == w + "tab")
                {
                    text.Append('\t');
                    tabCount++;
                    continue;
                }

                if (node.Name == w + "t")
                {
                    text.Append(node.Value);
                }
            }

            return new ParagraphTextSnapshot
            {
                Text = NormalizePlainText(text.ToString()),
                TabCount = tabCount
            };
        }

        private static SpecialConditionEvalOutcome EvaluateInsertedImage(
            ImageInsertConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("ChÆ°a cáº¥u hÃ¬nh Inserted Image (config trá»‘ng).");
            }

            if (string.IsNullOrWhiteSpace(config.ImageHash))
            {
                return Fail("ChÆ°a cÃ³ imageHash â€” áº£nh chuáº©n chÆ°a Ä‘Æ°á»£c upload/táº¡o hash á»Ÿ BE.");
            }

            var documentPart = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : NormalizeSourceFile(config.SourceFile);
            var documentRelsPath = string.IsNullOrWhiteSpace(config.RelsFile)
                ? "word/_rels/document.xml.rels"
                : NormalizeSourceFile(config.RelsFile);

            if (!package.TryGetXmlDocument(documentPart, out var documentDocument, out var documentError))
            {
                return Fail($"KhÃ´ng tÃ¬m tháº¥y {documentPart} trong file há»c sinh.");
            }

            if (!package.TryGetRelationships(documentRelsPath, out var relationships, out var relsError))
            {
                return Fail("KhÃ´ng tÃ¬m tháº¥y word/_rels/document.xml.rels.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            try
            {
                var expectedHash = ImageHashUtility.NormalizeHash(config.ImageHash);
                var expectedPerceptualHash = config.PerceptualHash;
                var expectedWrap = string.IsNullOrWhiteSpace(config.WrapType) ? null : config.WrapType.Trim();

                var paragraphs = documentDocument.Descendants(w + "body").Elements(w + "p").ToList();
                var drawings = paragraphs
                    .SelectMany((paragraph, paragraphIndex) => paragraph
                        .Descendants(w + "drawing")
                        .Select(drawing => new InsertedImageCandidate
                        {
                            Drawing = drawing,
                            ParagraphIndex = paragraphIndex
                        }))
                    .ToList();

                if (drawings.Count == 0)
                {
                    return Fail("KhÃ´ng tÃ¬m tháº¥y hÃ¬nh áº£nh (w:drawing) nÃ o trong tÃ i liá»‡u.");
                }

                string? lastMismatchInfo = null;

                foreach (var candidate in drawings)
                {
                    var drawing = candidate.Drawing;
                    var relationshipId = drawing.Descendants(a + "blip").FirstOrDefault()?.Attribute(r + "embed")?.Value;

                    if (string.IsNullOrWhiteSpace(relationshipId))
                    {
                        continue;
                    }

                    if (!relationships.TryGetValue(relationshipId, out var target) || string.IsNullOrWhiteSpace(target))
                    {
                        lastMismatchInfo = $"Relationship {relationshipId} khÃ´ng cÃ³ Target há»£p lá»‡.";
                        continue;
                    }

                    var imagePath = ResolveRelationshipTarget(documentPart, target);

                    if (!package.BinaryParts.TryGetValue(imagePath, out var imageBytes))
                    {
                        lastMismatchInfo = $"KhÃ´ng Ä‘á»c Ä‘Æ°á»£c áº£nh {imagePath} trong file há»c sinh.";
                        continue;
                    }

                    if (!IsImageMatch(package, imagePath, imageBytes, expectedHash, expectedPerceptualHash))
                    {
                        lastMismatchInfo = $"TÃ¬m tháº¥y áº£nh {imagePath} nhÆ°ng khÃ´ng Ä‘Ãºng ná»™i dung yÃªu cáº§u.";
                        continue;
                    }

                    var wrapMismatch = GetInsertedImageWrapMismatch(drawing, wp, expectedWrap);
                    if (wrapMismatch != null)
                    {
                        lastMismatchInfo = wrapMismatch;
                        continue;
                    }

                    var positionMismatch = GetInsertedImagePositionMismatch(
                        paragraphs,
                        w,
                        candidate.ParagraphIndex,
                        config.PositionConfig);
                    if (positionMismatch != null)
                    {
                        lastMismatchInfo = positionMismatch;
                        continue;
                    }

                    var sizeMismatch = GetInsertedImageSizeMismatch(drawing, wp, config.SizeConfig);
                    if (sizeMismatch != null)
                    {
                        lastMismatchInfo = sizeMismatch;
                        continue;
                    }

                    if (wrapMismatch == null)
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = BuildInsertedImageSuccessMessage(expectedWrap, config.PositionConfig, config.SizeConfig)
                        };
                    }

                    if (sizeMismatch == null)
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = BuildInsertedImageSuccessMessage(expectedWrap, config.PositionConfig, config.SizeConfig)
                        };
                    }

                    if (expectedWrap == null)
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = "ÄÃ£ chÃ¨n Ä‘Ãºng hÃ¬nh áº£nh yÃªu cáº§u."
                        };
                    }

                    var actualWrap = DetectWrapType(drawing, wp);

                    if (string.Equals(actualWrap, expectedWrap, StringComparison.OrdinalIgnoreCase))
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = $"ÄÃ£ chÃ¨n Ä‘Ãºng hÃ¬nh áº£nh yÃªu cáº§u vá»›i cháº¿ Ä‘á»™ ngáº¯t dÃ²ng '{expectedWrap}'."
                        };
                    }

                    lastMismatchInfo = $"ÄÃ£ tÃ¬m tháº¥y Ä‘Ãºng áº£nh yÃªu cáº§u nhÆ°ng cháº¿ Ä‘á»™ ngáº¯t dÃ²ng Ä‘ang lÃ  '{actualWrap}' thay vÃ¬ '{expectedWrap}'.";
                }

                return Fail(lastMismatchInfo ?? "KhÃ´ng tÃ¬m tháº¥y áº£nh nÃ o khá»›p vá»›i yÃªu cáº§u trong tÃ i liá»‡u.");
            }
            catch (XmlException ex)
            {
                return Fail($"KhÃ´ng thá»ƒ phÃ¢n tÃ­ch XML: {ex.Message}");
            }
        }

        /// <summary>
        /// wp:inline = "inline". wp:anchor cÃ³ 1 trong cÃ¡c pháº§n tá»­ wrap con:
        /// wrapSquare/wrapTight/wrapThrough/wrapTopAndBottom/wrapNone
        /// (behind/inFront phÃ¢n biá»‡t báº±ng attribute behindDoc trÃªn wp:anchor).
        /// </summary>
        private static string? GetInsertedImageWrapMismatch(XElement drawing, XNamespace wp, string? expectedWrap)
        {
            if (string.IsNullOrWhiteSpace(expectedWrap))
            {
                return null;
            }

            var actualWrap = DetectWrapType(drawing, wp);
            return string.Equals(actualWrap, expectedWrap, StringComparison.OrdinalIgnoreCase)
                ? null
                : $"Da tim thay dung anh nhung wrap la '{actualWrap}' thay vi '{expectedWrap}'.";
        }

        private static string? GetInsertedImagePositionMismatch(
            IReadOnlyList<XElement> paragraphs,
            XNamespace w,
            int imageParagraphIndex,
            ImagePositionConfig? config)
        {
            if (config == null || config.RequireBetween == false)
            {
                return null;
            }

            var afterText = NormalizePlainText(config.AfterText);
            var beforeText = NormalizePlainText(config.BeforeText);

            if (string.IsNullOrWhiteSpace(afterText) && string.IsNullOrWhiteSpace(beforeText))
            {
                return null;
            }

            var comparison = config.CaseSensitive == true
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            var paragraphTexts = paragraphs
                .Select(paragraph => BuildParagraphTextSnapshot(paragraph, w, excludeTextBoxContent: true).Text)
                .ToList();

            if (!string.IsNullOrWhiteSpace(afterText))
            {
                var afterIndex = paragraphTexts.FindLastIndex(
                    Math.Max(0, imageParagraphIndex),
                    text => text.Contains(afterText, comparison));

                if (afterIndex < 0)
                {
                    return $"Khong tim thay moc afterText '{afterText}' truoc anh.";
                }

                if (afterIndex >= imageParagraphIndex)
                {
                    return $"Anh khong nam sau afterText '{afterText}'.";
                }
            }

            if (!string.IsNullOrWhiteSpace(beforeText))
            {
                var beforeIndex = paragraphTexts.FindIndex(
                    imageParagraphIndex + 1,
                    text => text.Contains(beforeText, comparison));

                if (beforeIndex < 0)
                {
                    return $"Khong tim thay moc beforeText '{beforeText}' sau anh.";
                }

                if (beforeIndex <= imageParagraphIndex)
                {
                    return $"Anh khong nam truoc beforeText '{beforeText}'.";
                }
            }

            return null;
        }

        private static string? GetInsertedImageSizeMismatch(
            XElement drawing,
            XNamespace wp,
            ImageSizeConfig? config)
        {
            if (config == null || (!config.ExpectedWidthEmu.HasValue && !config.ExpectedHeightEmu.HasValue))
            {
                return null;
            }

            var extent = drawing.Element(wp + "inline")?.Element(wp + "extent")
                ?? drawing.Element(wp + "anchor")?.Element(wp + "extent");
            if (extent == null)
            {
                return "Khong doc duoc kich thuoc anh (wp:extent).";
            }

            var actualWidth = ParseLongAttribute(extent, "cx");
            var actualHeight = ParseLongAttribute(extent, "cy");
            var tolerance = Math.Max(0, config.ToleranceEmu ?? 0);

            if (config.ExpectedWidthEmu.HasValue)
            {
                if (!actualWidth.HasValue)
                {
                    return "Khong doc duoc chieu rong anh.";
                }

                if (Math.Abs(actualWidth.Value - config.ExpectedWidthEmu.Value) > tolerance)
                {
                    return $"Chieu rong anh la {actualWidth.Value} EMU, can {config.ExpectedWidthEmu.Value} EMU (+/- {tolerance}).";
                }
            }

            if (config.ExpectedHeightEmu.HasValue)
            {
                if (!actualHeight.HasValue)
                {
                    return "Khong doc duoc chieu cao anh.";
                }

                if (Math.Abs(actualHeight.Value - config.ExpectedHeightEmu.Value) > tolerance)
                {
                    return $"Chieu cao anh la {actualHeight.Value} EMU, can {config.ExpectedHeightEmu.Value} EMU (+/- {tolerance}).";
                }
            }

            return null;
        }

        private static string BuildInsertedImageSuccessMessage(
            string? expectedWrap,
            ImagePositionConfig? positionConfig,
            ImageSizeConfig? sizeConfig)
        {
            var checks = new List<string> { "dung anh" };

            if (!string.IsNullOrWhiteSpace(expectedWrap))
            {
                checks.Add($"wrap {expectedWrap.Trim()}");
            }

            if (positionConfig?.RequireBetween != false
                && (!string.IsNullOrWhiteSpace(positionConfig?.AfterText) || !string.IsNullOrWhiteSpace(positionConfig?.BeforeText)))
            {
                checks.Add("dung vi tri");
            }

            if (sizeConfig != null && (sizeConfig.ExpectedWidthEmu.HasValue || sizeConfig.ExpectedHeightEmu.HasValue))
            {
                checks.Add("dung kich thuoc");
            }

            return $"Da chen hinh anh {string.Join(", ", checks)}.";
        }

        private static string DetectWrapType(XElement drawing, XNamespace wp)
        {
            if (drawing.Element(wp + "inline") != null)
            {
                return ImageWrapTypes.Inline;
            }

            var anchor = drawing.Element(wp + "anchor");
            if (anchor == null)
            {
                return "unknown";
            }

            if (anchor.Element(wp + "wrapSquare") != null) return ImageWrapTypes.Square;
            if (anchor.Element(wp + "wrapTight") != null) return ImageWrapTypes.Tight;
            if (anchor.Element(wp + "wrapThrough") != null) return ImageWrapTypes.Through;
            if (anchor.Element(wp + "wrapTopAndBottom") != null) return ImageWrapTypes.TopAndBottom;

            if (anchor.Element(wp + "wrapNone") != null)
            {
                var behindDoc = anchor.Attribute("behindDoc")?.Value;
                return behindDoc == "1" ? ImageWrapTypes.Behind : ImageWrapTypes.InFront;
            }

            return "unknown";
        }

        private static SpecialConditionEvalOutcome EvaluatePictureBullet(
            PictureBulletConfig? config,
            OfficePackage package)
        {
            static SpecialConditionEvalOutcome Fail(string message) => new()
            {
                IsPassed = false,
                Message = message
            };

            if (config == null)
            {
                return Fail("ChÆ°a cáº¥u hÃ¬nh Picture Bullet (config trá»‘ng).");
            }

            if (string.IsNullOrWhiteSpace(config.ImageHash))
            {
                return Fail("ChÆ°a cÃ³ imageHash â€” áº£nh bullet chuáº©n chÆ°a Ä‘Æ°á»£c upload/táº¡o hash á»Ÿ BE.");
            }

            const string documentPart = "word/document.xml";
            const string numberingPath = "word/numbering.xml";
            const string numberingRelsPath = "word/_rels/numbering.xml.rels";

            if (!package.TryGetXmlDocument(documentPart, out var documentDocument, out var documentError))
            {
                return Fail($"KhÃ´ng tÃ¬m tháº¥y {documentPart} trong file há»c sinh.");
            }

            if (!package.TryGetXmlDocument(numberingPath, out var numberingDocument, out var numberingError))
            {
                return Fail("File há»c sinh khÃ´ng cÃ³ word/numbering.xml.");
            }

            if (!package.TryGetRelationships(numberingRelsPath, out var relationships, out var relsError))
            {
                return Fail("KhÃ´ng tÃ¬m tháº¥y word/_rels/numbering.xml.rels.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            try
            {
                var expectedHash = ImageHashUtility.NormalizeHash(config.ImageHash);
                var expectedPerceptualHash = config.PerceptualHash;

                // Duyá»‡t toÃ n bá»™ paragraph trong document.xml, tÃ¬m báº¥t ká»³ paragraph nÃ o
                // dÃ¹ng picture bullet á»Ÿ Ä‘Ãºng Level (náº¿u cÃ³ chá»‰ Ä‘á»‹nh) mÃ  áº£nh khá»›p expectedHash.
                var allParagraphs = documentDocument.Descendants(w + "p").ToList();
                var numById = numberingDocument.Descendants(w + "num")
                    .Where(element => !string.IsNullOrWhiteSpace(element.Attribute(w + "numId")?.Value))
                    .GroupBy(element => element.Attribute(w + "numId")!.Value, StringComparer.Ordinal)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
                var abstractNumById = numberingDocument.Descendants(w + "abstractNum")
                    .Where(element => !string.IsNullOrWhiteSpace(element.Attribute(w + "abstractNumId")?.Value))
                    .GroupBy(element => element.Attribute(w + "abstractNumId")!.Value, StringComparer.Ordinal)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
                var numPicBulletById = numberingDocument.Descendants(w + "numPicBullet")
                    .Where(element => !string.IsNullOrWhiteSpace(element.Attribute(w + "numPicBulletId")?.Value))
                    .GroupBy(element => element.Attribute(w + "numPicBulletId")!.Value, StringComparer.Ordinal)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

                var checkedAnyBulletParagraph = false;
                string? lastMismatchInfo = null;

                foreach (var paragraph in allParagraphs)
                {
                    var numPr = paragraph.Element(w + "pPr")?.Element(w + "numPr");
                    if (numPr == null)
                    {
                        continue;
                    }

                    var numIdVal = numPr.Element(w + "numId")?.Attribute(w + "val")?.Value;
                    if (!int.TryParse(numIdVal, out var numId))
                    {
                        continue;
                    }

                    var ilvlVal = numPr.Element(w + "ilvl")?.Attribute(w + "val")?.Value;
                    var ilvl = int.TryParse(ilvlVal, out var parsedIlvl) ? parsedIlvl : 0;

                    if (config.Level.HasValue && ilvl != config.Level.Value)
                    {
                        continue;
                    }

                    // numId -> w:num -> abstractNumId
                    if (!numById.TryGetValue(numId.ToString(), out var numElement))
                    {
                        continue;
                    }

                    var abstractNumIdVal = numElement.Element(w + "abstractNumId")?.Attribute(w + "val")?.Value;
                    if (!int.TryParse(abstractNumIdVal, out var abstractNumId))
                    {
                        continue;
                    }

                    // abstractNumId -> abstractNum -> lvl[ilvl] -> lvlPicBulletId
                    if (!abstractNumById.TryGetValue(abstractNumId.ToString(), out var abstractNumElement))
                    {
                        continue;
                    }

                    var lvlElement = abstractNumElement
                        .Elements(w + "lvl")
                        .FirstOrDefault(e => e.Attribute(w + "ilvl")?.Value == ilvl.ToString());

                    var lvlPicBulletIdVal = lvlElement?.Element(w + "lvlPicBulletId")?.Attribute(w + "val")?.Value;
                    if (!int.TryParse(lvlPicBulletIdVal, out var lvlPicBulletId))
                    {
                        // Level nÃ y khÃ´ng dÃ¹ng picture bullet -> khÃ´ng tÃ­nh lÃ  "Ä‘Ã£ kiá»ƒm tra bullet"
                        continue;
                    }

                    checkedAnyBulletParagraph = true;

                    // lvlPicBulletId -> numPicBullet -> r:id áº£nh
                    if (!numPicBulletById.TryGetValue(lvlPicBulletId.ToString(), out var numPicBulletElement))
                    {
                        lastMismatchInfo = $"numPicBulletId={lvlPicBulletId} khÃ´ng tá»“n táº¡i trong numbering.xml.";
                        continue;
                    }

                    var relationshipId =
                        numPicBulletElement.Descendants(v + "imagedata").FirstOrDefault()?.Attribute(r + "id")?.Value
                        ?? numPicBulletElement.Descendants(a + "blip").FirstOrDefault()?.Attribute(r + "embed")?.Value;

                    if (string.IsNullOrWhiteSpace(relationshipId))
                    {
                        lastMismatchInfo = $"numPicBulletId={lvlPicBulletId} khÃ´ng cÃ³ tham chiáº¿u áº£nh.";
                        continue;
                    }

                    if (!relationships.TryGetValue(relationshipId, out var target) || string.IsNullOrWhiteSpace(target))
                    {
                        lastMismatchInfo = $"Relationship {relationshipId} khÃ´ng cÃ³ Target há»£p lá»‡.";
                        continue;
                    }

                    var imagePath = ResolveRelationshipTarget(numberingPath, target);

                    if (!package.BinaryParts.TryGetValue(imagePath, out var imageBytes))
                    {
                        lastMismatchInfo = $"KhÃ´ng Ä‘á»c Ä‘Æ°á»£c áº£nh {imagePath} trong file há»c sinh.";
                        continue;
                    }

                    if (IsImageMatch(package, imagePath, imageBytes, expectedHash, expectedPerceptualHash))
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = "Picture bullet Ä‘Ãºng hÃ¬nh áº£nh yÃªu cáº§u."
                        };
                    }

                    lastMismatchInfo = $"TÃ¬m tháº¥y picture bullet (numId={numId}, level={ilvl}) nhÆ°ng áº£nh khÃ´ng khá»›p.";
                }

                if (!checkedAnyBulletParagraph)
                {
                    return Fail(config.Level.HasValue
                        ? $"KhÃ´ng tÃ¬m tháº¥y paragraph nÃ o dÃ¹ng picture bullet á»Ÿ level {config.Level.Value}."
                        : "KhÃ´ng tÃ¬m tháº¥y paragraph nÃ o dÃ¹ng picture bullet trong tÃ i liá»‡u.");
                }

                return Fail(lastMismatchInfo ?? "Picture bullet khÃ´ng Ä‘Ãºng hÃ¬nh áº£nh yÃªu cáº§u.");
            }
            catch (XmlException ex)
            {
                return Fail($"KhÃ´ng thá»ƒ phÃ¢n tÃ­ch XML: {ex.Message}");
            }
        }

        private static string ResolveRelationshipTarget(string sourcePart, string target)
        {
            target = target.Replace('\\', '/').Trim();

            if (target.StartsWith("/"))
            {
                return target.TrimStart('/');
            }

            var sourceDirectory = sourcePart.Contains('/')
                ? sourcePart[..sourcePart.LastIndexOf('/')]
                : string.Empty;

            var combined = string.IsNullOrWhiteSpace(sourceDirectory)
                ? target
                : $"{sourceDirectory}/{target}";

            var segments = combined.Split('/', StringSplitOptions.RemoveEmptyEntries);
            var stack = new Stack<string>();

            foreach (var segment in segments)
            {
                if (segment == ".") continue;
                if (segment == "..")
                {
                    if (stack.Count > 0) stack.Pop();
                    continue;
                }
                stack.Push(segment);
            }

            return string.Join("/", stack.Reverse());
        }

        private static ExpectedMatchResult MatchExpected(
            string sourceFile,
            string actualXml,
            string expectedValue,
            string compareMode,
            IReadOnlyList<string> ignoreAttributes,
            int? minOccurrences,
            int? maxOccurrences,
            XmlEvaluationCache cache)
        {
            var mode = string.IsNullOrWhiteSpace(compareMode)
                ? XmlGradingCompareModes.XmlContainsNormalized
                : compareMode.Trim();

            if (string.Equals(
                mode,
                XmlGradingCompareModes.ExactStringContains,
                StringComparison.OrdinalIgnoreCase))
            {
                return RawContains(
                    actualXml,
                    expectedValue,
                    trim: false);
            }

            if (string.Equals(
                mode,
                XmlGradingCompareModes.XmlContains,
                StringComparison.OrdinalIgnoreCase))
            {
                return RawContains(
                    actualXml,
                    expectedValue,
                    trim: true);
            }

            if (string.Equals(
                mode,
                XmlGradingCompareModes.XmlEquivalentWholeFile,
                StringComparison.OrdinalIgnoreCase))
            {
                return XmlEquivalentWholeFile(
                    actualXml,
                    expectedValue,
                    ignoreAttributes);
            }

            if (string.Equals(
                mode,
                XmlGradingCompareModes.XmlMinOccurrences,
                StringComparison.OrdinalIgnoreCase))
            {
                return XmlMinOccurrences(
                    sourceFile,
                    actualXml,
                    expectedValue,
                    ignoreAttributes,
                    minOccurrences,
                    maxOccurrences,
                    cache);
            }

            if (string.Equals(
                mode,
                XmlGradingCompareModes.XmlContainsNormalized,
                StringComparison.OrdinalIgnoreCase))
            {
                return XmlContainsNormalized(
                    sourceFile,
                    actualXml,
                    expectedValue,
                    ignoreAttributes,
                    cache);
            }

            // KhÃ´ng nÃªn Ã¢m tháº§m coi mode láº¡ lÃ  normalized
            return new ExpectedMatchResult
            {
                ExpectedValue = expectedValue,
                IsMatched = false,
                MatchIndex = null
            };
        }
        private static ExpectedMatchResult RawContains(string actualXml, string expectedValue, bool trim)
        {
            var expected = trim
                ? expectedValue.Trim()
                : expectedValue;

            var index = actualXml.IndexOf(
                expected,
                StringComparison.Ordinal);

            return new ExpectedMatchResult
            {
                ExpectedValue = expectedValue,
                IsMatched = index >= 0,
                MatchIndex = index >= 0 ? index : null
            };
        }

        private static List<ExpectedMatchResult> MatchExpectedOrdered(
            string sourceFile,
            string actualXml,
            IReadOnlyList<string> expectedValues,
            string compareMode,
            IReadOnlyList<string> ignoreAttributes,
            XmlEvaluationCache cache)
        {
            var mode = string.IsNullOrWhiteSpace(compareMode)
                ? XmlGradingCompareModes.XmlContainsNormalized
                : compareMode.Trim();

            // Chuáº©n hÃ³a search space 1 láº§n duy nháº¥t (khÃ´ng Ä‘á»•i thá»© tá»± kÃ½ tá»± nÃªn cursor váº«n há»£p lá»‡)
            var searchSpace = string.Equals(mode, XmlGradingCompareModes.XmlContainsNormalized, StringComparison.OrdinalIgnoreCase)
                ? cache.GetNormalizedActual(sourceFile, actualXml, ignoreAttributes)
                : actualXml;

            var results = new List<ExpectedMatchResult>();
            var cursor = 0;

            foreach (var expectedValue in expectedValues)
            {
                string expected;
                if (string.Equals(mode, XmlGradingCompareModes.XmlContainsNormalized, StringComparison.OrdinalIgnoreCase))
                {
                    expected = cache.GetNormalizedExpected(expectedValue, ignoreAttributes);
                }
                else if (string.Equals(mode, XmlGradingCompareModes.XmlContains, StringComparison.OrdinalIgnoreCase))
                {
                    expected = expectedValue.Trim();
                }
                else
                {
                    // exactStringContains
                    expected = expectedValue;
                }

                // TÃ¬m báº¯t Ä‘áº§u tá»« cursor hiá»‡n táº¡i -> Ä‘áº£m báº£o Ä‘Ãºng thá»© tá»± vÃ  liá»n máº¡ch,
                // khÃ´ng cho phÃ©p match á»Ÿ má»™t occurrence Ä‘á»©ng trÆ°á»›c fragment trÆ°á»›c Ä‘Ã³.
                var index = string.IsNullOrEmpty(expected)
                    ? -1
                    : searchSpace.IndexOf(expected, cursor, StringComparison.Ordinal);

                results.Add(new ExpectedMatchResult
                {
                    ExpectedValue = expectedValue,
                    IsMatched = index >= 0,
                    MatchIndex = index >= 0 ? index : null
                });

                if (index >= 0)
                {
                    cursor = index + expected.Length;
                }
                // Náº¿u khÃ´ng match, giá»¯ nguyÃªn cursor -> cÃ¡c fragment sau váº«n Ä‘Æ°á»£c thá»­,
                // nhÆ°ng IsPassed cuá»‘i cÃ¹ng sáº½ = false vÃ¬ cÃ³ Ã­t nháº¥t 1 fragment missing.
            }

            return results;
        }
        private static ExpectedMatchResult XmlContainsNormalized(
            string sourceFile,
            string actualXml,
            string expectedValue,
            IReadOnlyList<string> ignoreAttributes,
            XmlEvaluationCache cache)
        {
            var normalizedActual =
                cache.GetNormalizedActual(sourceFile, actualXml, ignoreAttributes);

            var normalizedExpected =
                cache.GetNormalizedExpected(expectedValue, ignoreAttributes);

            var index = normalizedActual.IndexOf(
                normalizedExpected,
                StringComparison.Ordinal);

            return new ExpectedMatchResult
            {
                ExpectedValue = expectedValue,
                IsMatched = index >= 0,
                MatchIndex = index >= 0 ? index : null
            };
        }

        private static ExpectedMatchResult XmlMinOccurrences(
            string sourceFile,
            string actualXml,
            string expectedValue,
            IReadOnlyList<string> ignoreAttributes,
            int? minOccurrences,
            int? maxOccurrences,
            XmlEvaluationCache cache)
        {
            var normalizedActual = cache.GetNormalizedActual(sourceFile, actualXml, ignoreAttributes);
            var normalizedExpected = cache.GetNormalizedExpected(expectedValue, ignoreAttributes);
            var firstIndex = string.IsNullOrEmpty(normalizedExpected)
                ? -1
                : normalizedActual.IndexOf(normalizedExpected, StringComparison.Ordinal);
            var occurrences = CountOccurrences(normalizedActual, normalizedExpected);
            var requiredOccurrences = Math.Max(1, minOccurrences ?? 1);
            var isWithinMax = !maxOccurrences.HasValue || occurrences <= maxOccurrences.Value;

            return new ExpectedMatchResult
            {
                ExpectedValue = expectedValue,
                IsMatched = occurrences >= requiredOccurrences && isWithinMax,
                MatchIndex = firstIndex >= 0 ? firstIndex : null
            };
        }

        private static int CountOccurrences(string value, string search)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(search))
            {
                return 0;
            }

            var count = 0;
            var index = 0;
            while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += search.Length;
            }

            return count;
        }

        private static string NormalizeXmlForComparison(
            string value,
            IReadOnlyList<string>? ignoreAttributes = null)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            if (ignoreAttributes?.Count > 0)
            {
                try
                {
                    return NormalizeXmlFragment(value, ignoreAttributes);
                }
                catch (XmlException)
                {
                    // Fall through to the lightweight whitespace normalization for partial/raw fragments.
                }
            }

            var text = value.Trim();

            // Chuáº©n hÃ³a khoáº£ng tráº¯ng
            text = Regex.Replace(
                text,
                @"\s+",
                " ");

            // Chuáº©n hÃ³a khoáº£ng tráº¯ng quanh dáº¥u =
            text = Regex.Replace(
                text,
                @"\s*=\s*",
                "=");

            // Chuáº©n hÃ³a khoáº£ng tráº¯ng trÆ°á»›c />
            text = Regex.Replace(
                text,
                @"\s*/>",
                "/>");

            return text.Trim();
        }
        private static ExpectedMatchResult XmlEquivalentWholeFile(
            string actualXml,
            string expectedValue,
            IReadOnlyList<string> ignoreAttributes)
        {
            try
            {
                var normalizedActual = NormalizeXmlFragment(actualXml, ignoreAttributes);
                var normalizedExpected = NormalizeXmlFragment(expectedValue, ignoreAttributes);
                var isMatched = string.Equals(normalizedActual, normalizedExpected, StringComparison.Ordinal);
                return new ExpectedMatchResult
                {
                    ExpectedValue = expectedValue,
                    IsMatched = isMatched,
                    MatchIndex = isMatched ? 0 : null
                };
            }
            catch (XmlException)
            {
                var isMatched = string.Equals(actualXml.Trim(), expectedValue.Trim(), StringComparison.Ordinal);
                return new ExpectedMatchResult
                {
                    ExpectedValue = expectedValue,
                    IsMatched = isMatched,
                    MatchIndex = isMatched ? 0 : null
                };
            }
        }

        private static string NormalizeXmlFragment(
            string xml,
            IReadOnlyList<string>? ignoreAttributes = null)
        {
            var wrapped = $"<__root {CommonOfficeNamespaceDeclarations}>{StripXmlDeclaration(xml)}</__root>";
            var document = XDocument.Parse(wrapped, LoadOptions.PreserveWhitespace);
            var normalized = string.Concat(document.Root!.Nodes().Select(node => NormalizeNode(node, ignoreAttributes)));
            return normalized;
        }

        private static string NormalizeNode(
            XNode node,
            IReadOnlyList<string>? ignoreAttributes = null)
        {
            return node switch
            {
                XElement element => NormalizeElement(element, ignoreAttributes),
                XCData cdata => SecurityElement.Escape(NormalizeText(cdata.Value)) ?? string.Empty,
                XText text => SecurityElement.Escape(NormalizeText(text.Value)) ?? string.Empty,
                _ => string.Empty
            };
        }

        private static string NormalizeElement(
            XElement element,
            IReadOnlyList<string>? ignoreAttributes = null)
        {
            var name = NormalizeName(element.Name);
            var attributes = element.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration)
                .Where(attribute => !ShouldIgnoreAttribute(attribute, ignoreAttributes))
                .OrderBy(attribute => NormalizeName(attribute.Name), StringComparer.Ordinal)
                .ThenBy(attribute => attribute.Value, StringComparer.Ordinal)
                .Select(attribute => $"{NormalizeName(attribute.Name)}=\"{SecurityElement.Escape(attribute.Value) ?? string.Empty}\"");

            var attributeText = string.Join(" ", attributes);
            var openTag = string.IsNullOrWhiteSpace(attributeText) ? $"<{name}>" : $"<{name} {attributeText}>";
            var children = string.Concat(element.Nodes().Select(node => NormalizeNode(node, ignoreAttributes)));

            return $"{openTag}{children}</{name}>";
        }

        private static bool ShouldIgnoreAttribute(
            XAttribute attribute,
            IReadOnlyList<string>? ignoreAttributes)
        {
            if (ignoreAttributes == null || ignoreAttributes.Count == 0)
            {
                return false;
            }

            var normalizedName = NormalizeName(attribute.Name);
            var localName = attribute.Name.LocalName;

            foreach (var ignoreAttribute in ignoreAttributes)
            {
                var token = NormalizeIgnoreAttributeToken(ignoreAttribute);
                if (string.IsNullOrWhiteSpace(token))
                {
                    continue;
                }

                if (token.EndsWith("*", StringComparison.Ordinal))
                {
                    var prefix = token[..^1];
                    if (localName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ||
                        normalizedName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    continue;
                }

                if (string.Equals(localName, token, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(normalizedName, token, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeIgnoreAttributeToken(string value)
        {
            var token = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                return string.Empty;
            }

            var atIndex = token.LastIndexOf('@');
            if (atIndex >= 0 && atIndex < token.Length - 1)
            {
                token = token[(atIndex + 1)..];
            }

            var colonIndex = token.LastIndexOf(':');
            if (colonIndex >= 0 && colonIndex < token.Length - 1)
            {
                token = token[(colonIndex + 1)..];
            }

            return token;
        }

        private static string NormalizeName(XName name)
        {
            return string.IsNullOrWhiteSpace(name.NamespaceName)
                ? name.LocalName
                : $"{{{name.NamespaceName}}}{name.LocalName}";
        }

        private static string NormalizeText(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : value.Trim();
        }

        private static string NormalizePlainText(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
            normalized = Regex.Replace(normalized, "[ ]+", " ");
            normalized = Regex.Replace(normalized, " *\t *", "\t");
            normalized = Regex.Replace(normalized, "\n+", "\n");
            return normalized;
        }

        private static string? NormalizeHexColor(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var normalized = value.Trim().TrimStart('#').ToUpperInvariant();
            return Regex.IsMatch(normalized, "^[0-9A-F]{6}$") ? normalized : value.Trim();
        }

        private static string? NormalizeBorderColor(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var normalized = value.Trim().TrimStart('#');
            var compact = Regex.Replace(normalized, "[\\s_-]+", string.Empty).ToLowerInvariant();

            return compact switch
            {
                "lightblue" => "00B0F0",
                "black" => "000000",
                "blue" => "0000FF",
                "red" => "FF0000",
                "green" => "00B050",
                "accent1" => "accent1",
                "accent2" => "accent2",
                "accent3" => "accent3",
                "accent4" => "accent4",
                "accent5" => "accent5",
                "accent6" => "accent6",
                _ => NormalizeHexColor(normalized)
            };
        }

        private static bool DoesBorderColorMatch(string actualColor, string expectedColor)
        {
            var actual = NormalizeBorderColor(actualColor);
            var expected = NormalizeBorderColor(expectedColor);

            if (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(expected, "00B0F0", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(actual, "accent1", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(actual, "5B9BD5", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(actual, "4F81BD", StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private static int? ParseIntAttribute(XElement element, string localName)
        {
            var value = element.Attributes().FirstOrDefault(attribute =>
                string.Equals(attribute.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase))?.Value;

            return int.TryParse(value, out var parsed) ? parsed : null;
        }

        private static long? ParseLongAttribute(XElement element, string localName)
        {
            var value = element.Attributes().FirstOrDefault(attribute =>
                string.Equals(attribute.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase))?.Value;

            return long.TryParse(value, out var parsed) ? parsed : null;
        }

        private static string? GetLineColor(XElement line, XNamespace a)
        {
            var solidFill = line.Element(a + "solidFill");
            var srgb = solidFill?.Element(a + "srgbClr")?.Attribute("val")?.Value;
            if (!string.IsNullOrWhiteSpace(srgb))
            {
                return NormalizeHexColor(srgb);
            }

            var scheme = solidFill?.Element(a + "schemeClr")?.Attribute("val")?.Value;
            return string.IsNullOrWhiteSpace(scheme) ? null : scheme.Trim();
        }

        private static bool DoesLineColorMatch(string? actualColor, string expectedColor)
        {
            if (string.IsNullOrWhiteSpace(actualColor))
            {
                return false;
            }

            var actual = NormalizeHexColor(actualColor);
            var expected = NormalizeHexColor(expectedColor);

            if (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(expected, "000000", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(actual, "tx1", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(actual, "dk1", StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private static bool HasVisiblePictureEffects(
            XElement picture,
            XElement shapeProperties,
            XNamespace a,
            XNamespace pic)
        {
            var effectElements = new[]
            {
                "outerShdw",
                "innerShdw",
                "prstShdw",
                "reflection",
                "glow",
                "softEdge",
                "scene3d",
                "sp3d"
            };

            if (shapeProperties.Descendants(a + "effectDag").Any())
            {
                return true;
            }

            if (effectElements.Any(name => shapeProperties.Descendants(a + name).Any()))
            {
                return true;
            }

            var styleEffectRefIndex = picture
                .Element(pic + "style")
                ?.Element(a + "effectRef")
                ?.Attribute("idx")
                ?.Value;

            return !string.IsNullOrWhiteSpace(styleEffectRefIndex)
                && !string.Equals(styleEffectRefIndex, "0", StringComparison.OrdinalIgnoreCase);
        }

        private static string StripXmlDeclaration(string xml)
        {
            var trimmed = xml.Trim();
            if (!trimmed.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase))
            {
                return trimmed;
            }

            var declarationEnd = trimmed.IndexOf("?>", StringComparison.Ordinal);
            return declarationEnd >= 0 ? trimmed[(declarationEnd + 2)..].Trim() : trimmed;
        }

        private static bool ApplyMatchPolicy(IReadOnlyList<ExpectedMatchResult> matches, string matchPolicy)
        {
            if (matches.Count == 0)
            {
                return false;
            }

            return matchPolicy switch
            {
                var value when string.Equals(value, XmlGradingMatchPolicies.Any, StringComparison.OrdinalIgnoreCase) =>
                    matches.Any(match => match.IsMatched),

                var value when string.Equals(value, XmlGradingMatchPolicies.Ordered, StringComparison.OrdinalIgnoreCase) =>
                    matches.All(match => match.IsMatched), // thá»© tá»± Ä‘Ã£ Ä‘Æ°á»£c Ä‘áº£m báº£o bá»Ÿi cursor khi tÃ¬m kiáº¿m

                _ => matches.All(match => match.IsMatched)
            };
        }

        private static XmlRuleValidationResult ValidateRuleSet(GradingRuleSet ruleSet)
        {
            var result = new XmlRuleValidationResult();

            if (string.IsNullOrWhiteSpace(ruleSet.Subject))
            {
                result.Errors.Add("subject khÃ´ng Ä‘Æ°á»£c rá»—ng.");
            }

            if (ruleSet.Projects.Count == 0)
            {
                result.Errors.Add("projects pháº£i cÃ³ Ã­t nháº¥t 1 project.");
            }

            foreach (var project in ruleSet.Projects)
            {
                var projectPrefix = string.IsNullOrWhiteSpace(project.ProjectCode) ? "project" : project.ProjectCode;
                if (string.IsNullOrWhiteSpace(project.ProjectCode))
                {
                    result.Errors.Add("project.projectCode khÃ´ng Ä‘Æ°á»£c rá»—ng.");
                }

                if (project.MaxScore <= 0)
                {
                    result.Errors.Add($"{projectPrefix}.maxScore pháº£i lá»›n hÆ¡n 0.");
                }

                foreach (var task in project.Tasks)
                {
                    var taskPrefix = string.IsNullOrWhiteSpace(task.TaskId) ? $"{projectPrefix}.task" : $"{projectPrefix}.{task.TaskId}";

                    if (string.IsNullOrWhiteSpace(task.TaskId))
                    {
                        result.Errors.Add($"{projectPrefix}.taskId khÃ´ng Ä‘Æ°á»£c rá»—ng.");
                    }

                    if (string.IsNullOrWhiteSpace(task.TaskName))
                    {
                        result.Errors.Add($"{taskPrefix}.taskName khÃ´ng Ä‘Æ°á»£c rá»—ng.");
                    }

                    if (task.MaxScore <= 0)
                    {
                        result.Errors.Add($"{taskPrefix}.maxScore pháº£i lá»›n hÆ¡n 0.");
                    }

                    var hasSpecialCondition = task.SpecialCondition != null
                        && !string.IsNullOrWhiteSpace(task.SpecialCondition.Type);

                    // Task há»£p lá»‡ khi cÃ³ Ã­t nháº¥t 1 Condition XML HOáº¶C cÃ³ specialCondition
                    // (khÃ´ng cÃ²n báº¯t buá»™c pháº£i cÃ³ Condition XML náº¿u Ä‘Ã£ dÃ¹ng specialCondition).
                    if (task.Conditions.Count == 0 && !hasSpecialCondition)
                    {
                        result.Errors.Add($"{taskPrefix}.conditions pháº£i cÃ³ Ã­t nháº¥t 1 condition, hoáº·c pháº£i cÃ³ specialCondition.");
                    }
                    else
                    {
                        // Tá»•ng Ä‘iá»ƒm = tá»•ng Conditions XML + Ä‘iá»ƒm riÃªng cá»§a specialCondition (náº¿u cÃ³).
                        // Cho phÃ©p Task chá»‰ dÃ¹ng specialCondition (0 Condition XML) hoáº·c káº¿t há»£p cáº£ hai.
                        var totalConditionScore = task.Conditions.Sum(condition => condition.Score)
                            + (hasSpecialCondition ? task.SpecialCondition!.Score : 0m);

                        if (totalConditionScore != task.MaxScore)
                        {
                            var scoreBreakdown = hasSpecialCondition
                                ? $"conditions + specialCondition = {totalConditionScore}"
                                : $"conditions = {totalConditionScore}";

                            result.Errors.Add($"{taskPrefix}: tá»•ng score ({scoreBreakdown}) pháº£i báº±ng task.maxScore ({task.MaxScore}).");
                        }
                    }

                    foreach (var condition in task.Conditions)
                    {
                        ValidateCondition(condition, taskPrefix, result);
                    }

                    if (hasSpecialCondition)
                    {
                        ValidateTaskSpecialCondition(task.SpecialCondition!, taskPrefix, result, ruleSet.Subject);
                    }
                }
            }

            return result;
        }

        private static void ValidateCondition(XmlGradingCondition condition, string taskPrefix, XmlRuleValidationResult result)
        {
            var conditionPrefix = string.IsNullOrWhiteSpace(condition.ConditionId)
                ? $"{taskPrefix}.condition"
                : $"{taskPrefix}.{condition.ConditionId}";

            if (string.IsNullOrWhiteSpace(condition.ConditionId))
            {
                result.Errors.Add($"{taskPrefix}.conditionId khÃ´ng Ä‘Æ°á»£c rá»—ng.");
            }

            if (condition.Score <= 0)
            {
                result.Errors.Add($"{conditionPrefix}.score pháº£i lá»›n hÆ¡n 0.");
            }

            if (string.IsNullOrWhiteSpace(condition.SourceFile))
            {
                result.Errors.Add($"{conditionPrefix}.sourceFile khÃ´ng Ä‘Æ°á»£c rá»—ng.");
            }
            else if (!IsSafeSourceFile(condition.SourceFile))
            {
                result.Errors.Add($"{conditionPrefix}.sourceFile khÃ´ng há»£p lá»‡ hoáº·c cÃ³ path traversal.");
            }

            if (condition.ExpectedVariants == null || condition.ExpectedVariants.Count == 0)
            {
                result.Errors.Add($"{conditionPrefix}.expectedVariants pháº£i cÃ³ Ã­t nháº¥t 1 variant.");
            }
            else
            {
                for (var variantIndex = 0; variantIndex < condition.ExpectedVariants.Count; variantIndex++)
                {
                    var variant = condition.ExpectedVariants[variantIndex];
                    if (variant == null || variant.ExpectedValues.Count == 0 || variant.ExpectedValues.Any(string.IsNullOrWhiteSpace))
                    {
                        result.Errors.Add($"{conditionPrefix}.expectedVariants[{variantIndex}].expectedValues pháº£i lÃ  array string khÃ´ng rá»—ng.");
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(condition.CompareMode))
            {
                condition.CompareMode = XmlGradingCompareModes.XmlContainsNormalized;
            }

            if (!XmlGradingCompareModes.Supported.Contains(condition.CompareMode))
            {
                result.Errors.Add($"{conditionPrefix}.compareMode khÃ´ng Ä‘Æ°á»£c há»— trá»£: {condition.CompareMode}.");
            }

            if (string.IsNullOrWhiteSpace(condition.MatchPolicy))
            {
                condition.MatchPolicy = XmlGradingMatchPolicies.All;
            }

            if (!XmlGradingMatchPolicies.Supported.Contains(condition.MatchPolicy))
            {
                result.Errors.Add($"{conditionPrefix}.matchPolicy khÃ´ng Ä‘Æ°á»£c há»— trá»£: {condition.MatchPolicy}.");
            }

            if (string.Equals(condition.CompareMode, XmlGradingCompareModes.XmlMinOccurrences, StringComparison.OrdinalIgnoreCase)
                && (!condition.MinOccurrences.HasValue || condition.MinOccurrences.Value <= 0))
            {
                result.Errors.Add($"{conditionPrefix}.minOccurrences phai lon hon 0 khi compareMode la xmlMinOccurrences.");
            }

            if (condition.MaxOccurrences.HasValue
                && condition.MinOccurrences.HasValue
                && condition.MaxOccurrences.Value < condition.MinOccurrences.Value)
            {
                result.Errors.Add($"{conditionPrefix}.maxOccurrences phai lon hon hoac bang minOccurrences.");
            }

            if (string.Equals(condition.CompareMode, XmlGradingCompareModes.XmlEquivalentWholeFile, StringComparison.OrdinalIgnoreCase) &&
                condition.ExpectedVariants?.Any(variant => variant.ExpectedValues.Count != 1) == true)
            {
                result.Errors.Add($"{conditionPrefix}.xmlEquivalentWholeFile chá»‰ há»— trá»£ Ä‘Ãºng 1 expectedValue trong má»—i expectedVariant.");
            }
        }

        private static void ValidateTaskSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result,
            string subject)
        {
            if (string.IsNullOrWhiteSpace(specialCondition.Type))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.type khÃ´ng Ä‘Æ°á»£c rá»—ng.");
                return;
            }

            if (!SpecialConditionTypes.Supported.Contains(specialCondition.Type))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.type khÃ´ng Ä‘Æ°á»£c há»— trá»£: {specialCondition.Type}.");
                return;
            }

            if (!IsSpecialConditionSupportedForSubject(specialCondition.Type, subject))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.type {specialCondition.Type} khong ho tro cho subject {NormalizeKey(subject)}.");
                return;
            }

            if (specialCondition.Score <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.score pháº£i lá»›n hÆ¡n 0.");
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureBullet, StringComparison.OrdinalIgnoreCase))
            {
                var config = specialCondition.Config;

                if (config == null)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.config khÃ´ng Ä‘Æ°á»£c null.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(config.ImageHash))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.config.imageHash khÃ´ng Ä‘Æ°á»£c rá»—ng (áº£nh bullet chuáº©n chÆ°a Ä‘Æ°á»£c upload/táº¡o hash).");
                }

                if (config.Level.HasValue && config.Level.Value < 0)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.config.level pháº£i >= 0.");
                }
            }

            // Má»šI
            if (string.Equals(specialCondition.Type, SpecialConditionTypes.InsertedImage, StringComparison.OrdinalIgnoreCase))
            {
                var config = specialCondition.ImageInsertConfig;

                if (config == null)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.imageInsertConfig khÃ´ng Ä‘Æ°á»£c null.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(config.ImageHash))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.imageInsertConfig.imageHash khÃ´ng Ä‘Æ°á»£c rá»—ng (áº£nh chuáº©n chÆ°a Ä‘Æ°á»£c upload/táº¡o hash).");
                }

                if (!string.IsNullOrWhiteSpace(config.WrapType) && !ImageWrapTypes.Supported.Contains(config.WrapType))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.imageInsertConfig.wrapType khÃ´ng Ä‘Æ°á»£c há»— trá»£: {config.WrapType}.");
                }
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ConvertTableToText, StringComparison.OrdinalIgnoreCase))
            {
                ValidateConvertTableToTextSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.Hyperlink, StringComparison.OrdinalIgnoreCase))
            {
                ValidateHyperlinkSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.SectionBreakBeforeText, StringComparison.OrdinalIgnoreCase))
            {
                ValidateSectionBreakBeforeTextSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureStyle, StringComparison.OrdinalIgnoreCase))
            {
                ValidatePictureStyleSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.TextBoxContainsText, StringComparison.OrdinalIgnoreCase))
            {
                ValidateTextBoxContainsTextSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PageMargins, StringComparison.OrdinalIgnoreCase))
            {
                ValidatePageMarginsSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.DocumentStyleSet, StringComparison.OrdinalIgnoreCase))
            {
                ValidateDocumentStyleSetSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PageBorder, StringComparison.OrdinalIgnoreCase))
            {
                ValidatePageBorderSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTableName, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelTableNameSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelWorksheetPageSetup, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelWorksheetPageSetupSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelClearCellFormatting, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelClearCellFormattingSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDataModelImport, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelDataModelImportSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelCompatibilityReport, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelCompatibilityReportSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelMergedRange, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelMergedRangeSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelCellHyperlink, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelCellHyperlinkSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelIconSetConditionalFormatting, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelIconSetConditionalFormattingSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartDataRange, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelChartDataRangeSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartStyle, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelChartStyleSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTextReplacement, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelTextReplacementSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelPrintTitles, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelPrintTitlesSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelNumberFormat, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelNumberFormatSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelChartLegend, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelChartLegendSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDefinedName, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelDefinedNameSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelFormulaReferences, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelFormulaReferencesSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelNoConditionalFormatting, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelNoConditionalFormattingSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelTextRotation, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelTextRotationSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelMultiColumnSort, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelMultiColumnSortSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelFreezePanes, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelFreezePanesSpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelDocumentProperty, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelDocumentPropertySpecialCondition(specialCondition, taskPrefix, result);
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.ExcelPrintArea, StringComparison.OrdinalIgnoreCase))
            {
                ValidateExcelPrintAreaSpecialCondition(specialCondition, taskPrefix, result);
            }
        }

        private static void ValidateExcelMergedRangeSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelMergedRangeConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelMergedRangeConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelMergedRangeConfig", result);

            if (!TryParseExcelRange(config.Range, out _, out _, out _, out _))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelMergedRangeConfig.range khong hop le.");
            }
        }

        private static void ValidateExcelCellHyperlinkSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelCellHyperlinkConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelCellHyperlinkConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelCellHyperlinkConfig", result);

            if (string.IsNullOrWhiteSpace(NormalizeExcelCellAddress(config.Cell)))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelCellHyperlinkConfig.cell khong hop le.");
            }

            if (string.IsNullOrWhiteSpace(config.Location) && string.IsNullOrWhiteSpace(config.Target))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelCellHyperlinkConfig phai co location hoac target.");
            }
        }

        private static void ValidateExcelIconSetConditionalFormattingSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelIconSetConditionalFormattingConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelIconSetConditionalFormattingConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelIconSetConditionalFormattingConfig", result);

            if (!TryParseExcelRange(config.Range, out _, out _, out _, out _))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelIconSetConditionalFormattingConfig.range khong hop le.");
            }

            if (string.IsNullOrWhiteSpace(config.IconSet))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelIconSetConditionalFormattingConfig.iconSet khong duoc rong.");
            }
        }

        private static void ValidateExcelChartDataRangeSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelChartDataRangeConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartDataRangeConfig khong duoc null.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(config.ChartSourceFile) && !IsSafeSourceFile(config.ChartSourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartDataRangeConfig.chartSourceFile khong hop le.");
            }

            if (string.IsNullOrWhiteSpace(config.ExpectedCategoryRange)
                && string.IsNullOrWhiteSpace(config.ExpectedValueRange)
                && (config.ExpectedValueRanges == null || config.ExpectedValueRanges.Count == 0)
                && (config.ExpectedSeriesNames == null || config.ExpectedSeriesNames.Count == 0)
                && !config.ExpectedPointCount.HasValue
                && string.IsNullOrWhiteSpace(config.ExpectedCategoryText))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartDataRangeConfig phai co it nhat 1 tieu chi can cham.");
            }

            if (config.ExpectedPointCount.HasValue && config.ExpectedPointCount.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartDataRangeConfig.expectedPointCount phai lon hon 0.");
            }
        }

        private static void ValidateExcelChartStyleSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelChartStyleConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartStyleConfig khong duoc null.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(config.ChartSourceFile) && !IsSafeSourceFile(config.ChartSourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartStyleConfig.chartSourceFile khong hop le.");
            }

            if (!string.IsNullOrWhiteSpace(config.StyleSourceFile) && !IsSafeSourceFile(config.StyleSourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartStyleConfig.styleSourceFile khong hop le.");
            }

            if (!config.StyleId.HasValue || config.StyleId.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartStyleConfig.styleId phai lon hon 0.");
            }
        }

        private static void ValidateExcelTextReplacementSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelTextReplacementConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextReplacementConfig khong duoc null.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(config.SourceFile) && !IsSafeSourceFile(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextReplacementConfig.sourceFile khong hop le.");
            }

            if (string.IsNullOrWhiteSpace(config.OldText))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextReplacementConfig.oldText khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(config.NewText))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextReplacementConfig.newText khong duoc rong.");
            }

            if (config.MinNewTextOccurrences.HasValue && config.MinNewTextOccurrences.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextReplacementConfig.minNewTextOccurrences phai lon hon 0.");
            }
        }

        private static void ValidateExcelPrintTitlesSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelPrintTitlesConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelPrintTitlesConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.WorksheetName))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelPrintTitlesConfig.worksheetName khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(NormalizeExcelPrintRows(config.ExpectedRows)))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelPrintTitlesConfig.expectedRows khong hop le. Vi du: 1:3.");
            }
        }

        private static void ValidateExcelNumberFormatSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelNumberFormatConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNumberFormatConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelNumberFormatConfig", result);

            if (string.IsNullOrWhiteSpace(NormalizeExcelRangeOrColumnRange(config.Range)))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNumberFormatConfig.range khong hop le. Vi du: B:E hoac B4:E20.");
            }

            var hasCategory = !string.IsNullOrWhiteSpace(config.Category);
            if (hasCategory && !IsSupportedExcelNumberFormatCategory(config.Category))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNumberFormatConfig.category khong hop le. Gia tri hop le: general, number, currency, accounting, percentage, date, time, custom.");
            }

            if (config.DecimalPlaces < 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNumberFormatConfig.decimalPlaces phai lon hon hoac bang 0.");
            }

            if (!hasCategory && (config.AllowedNumberFormatIds == null || config.AllowedNumberFormatIds.Count == 0))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNumberFormatConfig.allowedNumberFormatIds phai co it nhat 1 id.");
            }
            else if (config.AllowedNumberFormatIds != null && config.AllowedNumberFormatIds.Any(value => value <= 0))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNumberFormatConfig.allowedNumberFormatIds phai lon hon 0.");
            }
        }

        private static bool IsSupportedExcelNumberFormatCategory(string? category)
        {
            return category?.Trim().ToLowerInvariant() is "general" or "number" or "currency" or "accounting" or "percentage" or "date" or "time" or "custom";
        }

        private static void ValidateExcelChartLegendSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelChartLegendConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartLegendConfig khong duoc null.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(config.ChartSourceFile) && !IsSafeSourceFile(config.ChartSourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartLegendConfig.chartSourceFile khong hop le.");
            }

            var supportedPositions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "t", "b", "l", "r", "tr" };
            var position = string.IsNullOrWhiteSpace(config.Position) ? "t" : config.Position.Trim();
            if (!supportedPositions.Contains(position))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelChartLegendConfig.position khong hop le. Dung t, b, l, r, hoac tr.");
            }
        }

        private static void ValidateExcelDefinedNameSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelDefinedNameConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDefinedNameConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.Name))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDefinedNameConfig.name khong duoc rong.");
            }

            if (config.ExpectedRanges == null || config.ExpectedRanges.Count == 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDefinedNameConfig.expectedRanges phai co it nhat 1 vung.");
            }
            else if (config.ExpectedRanges.Any(range => string.IsNullOrWhiteSpace(NormalizeExcelFormulaReference(range))))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDefinedNameConfig.expectedRanges co vung khong hop le.");
            }
        }

        private static void ValidateExcelFormulaReferencesSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelFormulaReferencesConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFormulaReferencesConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelFormulaReferencesConfig", result);

            if (string.IsNullOrWhiteSpace(NormalizeExcelCellAddress(config.Cell)))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFormulaReferencesConfig.cell khong hop le.");
            }

            if ((config.RequiredReferences == null || config.RequiredReferences.Count == 0)
                && (config.RequiredFunctions == null || config.RequiredFunctions.Count == 0)
                && (config.RequiredFormulaFragments == null || config.RequiredFormulaFragments.Count == 0)
                && string.IsNullOrWhiteSpace(config.ExpectedFormula)
                && string.IsNullOrWhiteSpace(config.ExpectedValue))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFormulaReferencesConfig phai co it nhat 1 tieu chi: requiredReferences, requiredFunctions, requiredFormulaFragments, expectedFormula, hoac expectedValue.");
            }

            if (!string.IsNullOrWhiteSpace(config.SourceFile) && !IsSafeSourceFile(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFormulaReferencesConfig.sourceFile khong hop le.");
            }
        }

        private static void ValidateExcelNoConditionalFormattingSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelNoConditionalFormattingConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNoConditionalFormattingConfig khong duoc null.");
                return;
            }

            if (config.RequireAllWorksheets != true)
            {
                ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelNoConditionalFormattingConfig", result);
            }
            else if (!string.IsNullOrWhiteSpace(config.SourceFile) && !IsSafeSourceFile(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelNoConditionalFormattingConfig.sourceFile khong hop le.");
            }
        }

        private static void ValidateExcelTextRotationSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelTextRotationConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextRotationConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelTextRotationConfig", result);

            if (config.ExpectedTexts == null || config.ExpectedTexts.Count == 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextRotationConfig.expectedTexts phai co it nhat 1 tieu de.");
            }

            if (config.AllowedTextRotationValues == null || config.AllowedTextRotationValues.Count == 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextRotationConfig.allowedTextRotationValues phai co it nhat 1 gia tri.");
            }
            else if (config.AllowedTextRotationValues.Any(value => value is < 0 or > 180))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTextRotationConfig.allowedTextRotationValues phai nam trong khoang 0-180.");
            }
        }

        private static void ValidateExcelMultiColumnSortSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelMultiColumnSortConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelMultiColumnSortConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelMultiColumnSortConfig", result);

            if (config.HeaderRow <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelMultiColumnSortConfig.headerRow phai lon hon 0.");
            }

            if (!string.IsNullOrWhiteSpace(config.DataRange) && !TryParseExcelRange(config.DataRange, out _, out _, out _, out _))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelMultiColumnSortConfig.dataRange khong hop le. Vi du: A5:H26.");
            }

            if (config.KeyColumns == null || config.KeyColumns.Count == 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelMultiColumnSortConfig.keyColumns phai co it nhat 1 cot sap xep.");
                return;
            }

            for (var index = 0; index < config.KeyColumns.Count; index++)
            {
                var key = config.KeyColumns[index];
                var hasHeaderName = !string.IsNullOrWhiteSpace(key.HeaderName);
                var hasColumn = Regex.IsMatch(key.Column ?? string.Empty, "^[A-Z]{1,3}$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!hasHeaderName && !hasColumn)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.excelMultiColumnSortConfig.keyColumns[{index}] phai co headerName hoac column hop le.");
                }
            }
        }

        private static void ValidateExcelFreezePanesSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelFreezePanesConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFreezePanesConfig khong duoc null.");
                return;
            }

            ValidateExcelWorksheetLocator(config.WorksheetName, config.SourceFile, $"{taskPrefix}.specialCondition.excelFreezePanesConfig", result);

            if (string.IsNullOrWhiteSpace(NormalizeExcelCellAddress(config.TopLeftCell)))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFreezePanesConfig.topLeftCell khong hop le.");
            }

            if (config.YSplit < 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFreezePanesConfig.ySplit phai lon hon hoac bang 0.");
            }

            if (config.XSplit < 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelFreezePanesConfig.xSplit phai lon hon hoac bang 0.");
            }
        }

        private static void ValidateExcelDocumentPropertySpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelDocumentPropertyConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDocumentPropertyConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.PropertyName))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDocumentPropertyConfig.propertyName khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(config.ExpectedValue))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDocumentPropertyConfig.expectedValue khong duoc rong.");
            }

            if (!string.IsNullOrWhiteSpace(config.SourceFile) && !IsSafeSourceFile(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDocumentPropertyConfig.sourceFile khong hop le.");
            }
        }

        private static void ValidateExcelPrintAreaSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelPrintAreaConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelPrintAreaConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.WorksheetName))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelPrintAreaConfig.worksheetName khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(NormalizeExcelFormulaReference(config.ExpectedRange)))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelPrintAreaConfig.expectedRange khong hop le. Vi du: A1:F17.");
            }
        }

        private static void ValidateExcelWorksheetLocator(
            string? worksheetName,
            string? sourceFile,
            string prefix,
            XmlRuleValidationResult result)
        {
            if (string.IsNullOrWhiteSpace(worksheetName) && string.IsNullOrWhiteSpace(sourceFile))
            {
                result.Errors.Add($"{prefix} phai co worksheetName hoac sourceFile.");
            }

            if (!string.IsNullOrWhiteSpace(sourceFile) && !IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{prefix}.sourceFile khong hop le.");
            }
        }

        private static void ValidateExcelTableNameSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelTableNameConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTableNameConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.ExpectedName))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTableNameConfig.expectedName khong duoc rong.");
            }

            if (!string.IsNullOrWhiteSpace(config.SourceFile) && !IsSafeSourceFile(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelTableNameConfig.sourceFile khong hop le.");
            }
        }

        private static void ValidateExcelWorksheetPageSetupSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelWorksheetPageSetupConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelWorksheetPageSetupConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.WorksheetName) && string.IsNullOrWhiteSpace(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelWorksheetPageSetupConfig phai co worksheetName hoac sourceFile.");
            }

            if (!string.IsNullOrWhiteSpace(config.SourceFile) && !IsSafeSourceFile(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelWorksheetPageSetupConfig.sourceFile khong hop le.");
            }

            var orientation = string.IsNullOrWhiteSpace(config.Orientation) ? "landscape" : config.Orientation.Trim();
            if (!string.Equals(orientation, "portrait", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(orientation, "landscape", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelWorksheetPageSetupConfig.orientation phai la portrait hoac landscape.");
            }
        }

        private static void ValidateExcelClearCellFormattingSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelClearCellFormattingConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelClearCellFormattingConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.WorksheetName) && string.IsNullOrWhiteSpace(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelClearCellFormattingConfig phai co worksheetName hoac sourceFile.");
            }

            if (!string.IsNullOrWhiteSpace(config.SourceFile) && !IsSafeSourceFile(config.SourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelClearCellFormattingConfig.sourceFile khong hop le.");
            }

            if (!TryParseExcelRange(config.Range, out _, out _, out _, out _))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelClearCellFormattingConfig.range khong hop le.");
            }

            if (config.DefaultStyleId.HasValue && config.DefaultStyleId.Value < 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelClearCellFormattingConfig.defaultStyleId phai >= 0.");
            }
        }

        private static void ValidateExcelDataModelImportSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelDataModelImportConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelDataModelImportConfig khong duoc null.");
                return;
            }

            if (config.RequireConnection == false
                && config.RequireDataModel == false
                && string.IsNullOrWhiteSpace(config.SourceFileName)
                && string.IsNullOrWhiteSpace(config.ExpectedConnectionName))
            {
                result.Warnings.Add($"{taskPrefix}.specialCondition.excelDataModelImportConfig dang qua rong.");
            }
        }

        private static void ValidateExcelCompatibilityReportSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ExcelCompatibilityReportConfig;
            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.excelCompatibilityReportConfig khong duoc null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(config.WorksheetName)
                && (config.ExpectedTexts == null || config.ExpectedTexts.All(string.IsNullOrWhiteSpace)))
            {
                result.Warnings.Add($"{taskPrefix}.specialCondition.excelCompatibilityReportConfig nen co worksheetName hoac expectedTexts.");
            }
        }

        private static void ValidateTextBoxContainsTextSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.TextBoxContainsTextConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.textBoxContainsTextConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : config.SourceFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.textBoxContainsTextConfig.sourceFile khong hop le.");
            }

            if (string.IsNullOrWhiteSpace(config.ExpectedText))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.textBoxContainsTextConfig.expectedText khong duoc rong.");
            }

            if (!string.IsNullOrWhiteSpace(config.MatchMode))
            {
                var supportedMatchModes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "exact",
                    "contains"
                };

                if (!supportedMatchModes.Contains(config.MatchMode.Trim()))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.textBoxContainsTextConfig.matchMode khong hop le.");
                }
            }

            if (config.TargetOccurrence.HasValue && config.TargetOccurrence.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.textBoxContainsTextConfig.targetOccurrence phai lon hon 0.");
            }
        }

        private static void ValidatePageMarginsSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.PageMarginsConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pageMarginsConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : config.SourceFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pageMarginsConfig.sourceFile khong hop le.");
            }

            if (!config.Top.HasValue && !config.Bottom.HasValue && !config.Left.HasValue && !config.Right.HasValue && !config.Gutter.HasValue)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pageMarginsConfig phai co it nhat 1 margin can cham.");
            }

            foreach (var (name, value) in new[]
            {
                ("top", config.Top),
                ("bottom", config.Bottom),
                ("left", config.Left),
                ("right", config.Right),
                ("gutter", config.Gutter)
            })
            {
                if (value.HasValue && value.Value < 0)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.pageMarginsConfig.{name} phai >= 0.");
                }
            }
        }

        private static void ValidateDocumentStyleSetSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.DocumentStyleSetConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.documentStyleSetConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/styles.xml"
                : config.SourceFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.documentStyleSetConfig.sourceFile khong hop le.");
            }

            if (config.ExpectedFragments == null || config.ExpectedFragments.Count == 0 || config.ExpectedFragments.All(string.IsNullOrWhiteSpace))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.documentStyleSetConfig.expectedFragments phai co it nhat 1 fragment.");
            }

            if (!string.IsNullOrWhiteSpace(config.MatchPolicy)
                && !XmlGradingMatchPolicies.Supported.Contains(config.MatchPolicy.Trim()))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.documentStyleSetConfig.matchPolicy khong duoc ho tro: {config.MatchPolicy}.");
            }
        }

        private static void ValidatePageBorderSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.PageBorderConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pageBorderConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : config.SourceFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pageBorderConfig.sourceFile khong hop le.");
            }

            if (config.RequiredWidth.HasValue && config.RequiredWidth.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pageBorderConfig.requiredWidth phai lon hon 0.");
            }

            if (config.MinWidth.HasValue && config.MinWidth.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pageBorderConfig.minWidth phai lon hon 0.");
            }

            if (string.IsNullOrWhiteSpace(config.RequiredStyle))
            {
                result.Warnings.Add($"{taskPrefix}.specialCondition.pageBorderConfig.requiredStyle dang trong nen se khong kiem tra kieu net.");
            }

            if (string.IsNullOrWhiteSpace(config.RequiredColor)
                && (config.AllowedColors == null || config.AllowedColors.Count == 0 || config.AllowedColors.All(string.IsNullOrWhiteSpace)))
            {
                result.Warnings.Add($"{taskPrefix}.specialCondition.pageBorderConfig khong co mau can cham nen se khong kiem tra mau border.");
            }
        }

        private static void ValidateSectionBreakBeforeTextSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.SectionBreakBeforeTextConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.sectionBreakBeforeTextConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : config.SourceFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.sectionBreakBeforeTextConfig.sourceFile khong hop le.");
            }

            if (string.IsNullOrWhiteSpace(config.TargetText))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.sectionBreakBeforeTextConfig.targetText khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(config.BreakType))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.sectionBreakBeforeTextConfig.breakType khong duoc rong.");
            }
            else
            {
                var supportedBreakTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "continuous",
                    "nextPage",
                    "evenPage",
                    "oddPage",
                    "nextColumn"
                };

                if (!supportedBreakTypes.Contains(config.BreakType.Trim()))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.sectionBreakBeforeTextConfig.breakType khong hop le.");
                }
            }

            if (config.TargetOccurrence.HasValue && config.TargetOccurrence.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.sectionBreakBeforeTextConfig.targetOccurrence phai lon hon 0.");
            }
        }

        private static void ValidatePictureStyleSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.PictureStyleConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pictureStyleConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : config.SourceFile;

            var relsFile = string.IsNullOrWhiteSpace(config.RelsFile)
                ? "word/_rels/document.xml.rels"
                : config.RelsFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pictureStyleConfig.sourceFile khong hop le.");
            }

            if (!IsSafeSourceFile(relsFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pictureStyleConfig.relsFile khong hop le.");
            }

            if (config.TargetImageIndex.HasValue && config.TargetImageIndex.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pictureStyleConfig.targetImageIndex phai lon hon 0.");
            }

            if (!string.IsNullOrWhiteSpace(config.StylePreset))
            {
                var supportedStylePresets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "simpleFrameBlack",
                    "custom"
                };

                if (!supportedStylePresets.Contains(config.StylePreset.Trim()))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.pictureStyleConfig.stylePreset khong hop le.");
                }
            }

            if (!string.IsNullOrWhiteSpace(config.RequiredLineColor)
                && !Regex.IsMatch(config.RequiredLineColor.Trim().TrimStart('#'), "^[0-9A-Fa-f]{6}$"))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pictureStyleConfig.requiredLineColor phai la ma mau hex 6 ky tu.");
            }

            if (config.MinLineWidth.HasValue && config.MinLineWidth.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.pictureStyleConfig.minLineWidth phai lon hon 0.");
            }

            if (string.IsNullOrWhiteSpace(config.ImageHash))
            {
                result.Warnings.Add($"{taskPrefix}.specialCondition.pictureStyleConfig.imageHash nen duoc cau hinh de tranh cham nham anh khi tai lieu co nhieu anh.");
            }
        }

        private static void ValidateHyperlinkSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.HyperlinkConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.hyperlinkConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : config.SourceFile;

            var relsFile = string.IsNullOrWhiteSpace(config.RelsFile)
                ? "word/_rels/document.xml.rels"
                : config.RelsFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.hyperlinkConfig.sourceFile khong hop le.");
            }

            if (!IsSafeSourceFile(relsFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.hyperlinkConfig.relsFile khong hop le.");
            }

            if (string.IsNullOrWhiteSpace(config.DisplayText))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.hyperlinkConfig.displayText khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(config.Url))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.hyperlinkConfig.url khong duoc rong.");
            }

            if (string.IsNullOrWhiteSpace(config.AnchorTextBefore))
            {
                result.Warnings.Add($"{taskPrefix}.specialCondition.hyperlinkConfig.anchorTextBefore nen duoc cau hinh neu displayText co the lap lai nhieu vi tri.");
            }
        }

        private static void ValidateConvertTableToTextSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
        {
            var config = specialCondition.ConvertTableToTextConfig;

            if (config == null)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.convertTableToTextConfig khong duoc null.");
                return;
            }

            var sourceFile = string.IsNullOrWhiteSpace(config.SourceFile)
                ? "word/document.xml"
                : config.SourceFile;

            if (!IsSafeSourceFile(sourceFile))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.convertTableToTextConfig.sourceFile khong hop le.");
            }

            if (config.MinRows.HasValue && config.MinRows.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.convertTableToTextConfig.minRows phai lon hon 0.");
            }

            if (config.MinTabsPerRow.HasValue && config.MinTabsPerRow.Value <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.convertTableToTextConfig.minTabsPerRow phai lon hon 0.");
            }

            if ((config.ExpectedRows == null || config.ExpectedRows.Count == 0) && !config.MinRows.HasValue)
            {
                result.Warnings.Add($"{taskPrefix}.specialCondition.convertTableToTextConfig nen co expectedRows hoac minRows de tranh dieu kien qua rong.");
            }
        }

        private static bool IsSafeSourceFile(string sourceFile)
        {
            var normalized = NormalizeSourceFile(sourceFile);

            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            // Pháº£i lÃ  Ä‘Æ°á»ng dáº«n tÆ°Æ¡ng Ä‘á»‘i trong Office ZIP package
            if (Path.IsPathRooted(normalized))
                return false;

            // KhÃ´ng Ä‘Æ°á»£c báº¯t Ä‘áº§u báº±ng /
            if (normalized.StartsWith("/", StringComparison.Ordinal))
                return false;

            // KhÃ´ng cho phÃ©p path traversal
            var segments = normalized.Split(
                '/',
                StringSplitOptions.RemoveEmptyEntries);

            if (segments.Any(segment =>
                segment == ".."))
            {
                return false;
            }

            // KhÃ´ng cho phÃ©p segment rá»—ng báº¥t thÆ°á»ng
            if (segments.Any(segment =>
                segment == "."))
            {
                return false;
            }

            // Chá»‰ cho phÃ©p XML hoáº·c RELS
            return normalized.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                || normalized.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
        }
        private static string NormalizeSourceFile(string sourceFile)
        {
            if (string.IsNullOrWhiteSpace(sourceFile))
                return string.Empty;

            var normalized = sourceFile
                .Trim()
                .Replace('\\', '/');

            // Loáº¡i bá» ./ á»Ÿ Ä‘áº§u
            while (normalized.StartsWith("./", StringComparison.Ordinal))
            {
                normalized = normalized[2..];
            }

            // KhÃ´ng cho phÃ©p / á»Ÿ Ä‘áº§u
            normalized = normalized.TrimStart('/');

            return normalized;
        }

        private static string NormalizeKey(string value)
        {
            return (value ?? string.Empty).Trim().ToLowerInvariant();
        }

        private static void ApplyProjectScoringModel(GradingResult result)
        {
            var gradableTasks = result.TaskResults
                .Where(task => task.MaxScore > 0m)
                .ToList();

            if (gradableTasks.Count == 0)
            {
                result.MaxScore = 0m;
                result.TotalScore = 0m;
                return;
            }

            // MaxScore cá»§a bÃ i = tá»•ng MaxScore do ngÆ°á»i táº¡o
            // cáº¥u hÃ¬nh cho tá»«ng Task.
            result.MaxScore = gradableTasks.Sum(task => task.MaxScore);

            // Äiá»ƒm thá»±c táº¿ = tá»•ng Ä‘iá»ƒm cÃ¡c Task Ä‘áº¡t Ä‘Æ°á»£c.
            result.TotalScore = Math.Round(
                gradableTasks.Sum(task => task.Score),
                2,
                MidpointRounding.AwayFromZero
            );

            // KhÃ´ng cho vÆ°á»£t quÃ¡ Ä‘iá»ƒm tá»‘i Ä‘a.
            result.TotalScore = Math.Min(
                result.TotalScore,
                result.MaxScore
            );
        }
        private static GradingRuleSet BuildProject22Task1RuleSet()
        {
            return new GradingRuleSet
            {
                Subject = "excel",
                Version = "2026.08.26",
                IsActive = true,
                Projects = new List<ProjectXmlRule>
                {
                    new()
                    {
                        ProjectCode = "project22",
                        ProjectName = "Excel Project 22",
                        MaxScore = StandardProjectMaxScore,
                        Tasks = new List<TaskXmlRule>
                        {
                            new()
                            {
                                TaskId = "P22-T1",
                                TaskName = "Sao chÃ©p Ä‘á»‹nh dáº¡ng tá»« tiÃªu Ä‘á» vÃ  phá»¥ Ä‘á» cá»§a trang tÃ­nh Task sang Project.",
                                MaxScore = 18m,
                                Conditions = new List<XmlGradingCondition>
                                {
                                    new()
                                    {
                                        ConditionId = "P22-T1-C01",
                                        Score = 2m,
                                        SourceFile = "xl/worksheets/sheet1.xml",
                                        ExpectedVariants = new List<XmlExpectedVariant>
                                        {
                                            new()
                                            {
                                                ExpectedValues = new List<string>
                                                {
                                                    "<c r=\"A1\" t=\"s\"><v>0</v></c>",
                                                    "<c r=\"A2\" t=\"s\"><v>1</v></c>"
                                                }
                                            }
                                        },
                                        CompareMode = XmlGradingCompareModes.XmlContainsNormalized,
                                        MatchPolicy = XmlGradingMatchPolicies.All,
                                        Feedback = new ConditionFeedback
                                        {
                                            SuccessDetail = "ÄÃ£ xÃ¡c nháº­n XML tiÃªu Ä‘á»/phá»¥ Ä‘á» nguá»“n táº¡i Task!A1:A2.",
                                            ErrorMessage = "KhÃ´ng tÃ¬m tháº¥y Ä‘áº§y Ä‘á»§ XML tiÃªu Ä‘á»/phá»¥ Ä‘á» nguá»“n táº¡i Task!A1:A2.",
                                            FixAction = "Kiá»ƒm tra láº¡i ná»™i dung tiÃªu Ä‘á»/phá»¥ Ä‘á» trong sheet Task."
                                        }
                                    },
                                    new()
                                    {
                                        ConditionId = "P22-T1-C02",
                                        Score = 4m,
                                        SourceFile = "xl/worksheets/sheet2.xml",
                                        ExpectedVariants = new List<XmlExpectedVariant>
                                        {
                                            new()
                                            {
                                                ExpectedValues = new List<string>
                                                {
                                                    "<c r=\"A1\" s=\"5\" t=\"s\"><v>0</v></c>"
                                                }
                                            }
                                        },
                                        CompareMode = XmlGradingCompareModes.XmlContainsNormalized,
                                        MatchPolicy = XmlGradingMatchPolicies.All,
                                        Feedback = new ConditionFeedback
                                        {
                                            SuccessDetail = "Project!A1 cÃ³ XML Ä‘á»‹nh dáº¡ng Ä‘Ãºng.",
                                            ErrorMessage = "Project!A1 chÆ°a cÃ³ XML Ä‘á»‹nh dáº¡ng Ä‘Ãºng.",
                                            FixAction = "DÃ¹ng Format Painter sao chÃ©p Ä‘á»‹nh dáº¡ng tá»« Task!A1 sang Project!A1."
                                        }
                                    },
                                    new()
                                    {
                                        ConditionId = "P22-T1-C03",
                                        Score = 6m,
                                        SourceFile = "xl/worksheets/sheet2.xml",
                                        ExpectedVariants = new List<XmlExpectedVariant>
                                        {
                                            new()
                                            {
                                                ExpectedValues = new List<string>
                                                {
                                                    "<c r=\"A2\" s=\"6\" t=\"s\"><v>1</v></c>"
                                                }
                                            }
                                        },
                                        CompareMode = XmlGradingCompareModes.XmlContainsNormalized,
                                        MatchPolicy = XmlGradingMatchPolicies.All,
                                        Feedback = new ConditionFeedback
                                        {
                                            SuccessDetail = "Project!A2 cÃ³ XML Ä‘á»‹nh dáº¡ng Ä‘Ãºng.",
                                            ErrorMessage = "Project!A2 chÆ°a cÃ³ XML Ä‘á»‹nh dáº¡ng Ä‘Ãºng.",
                                            FixAction = "DÃ¹ng Format Painter sao chÃ©p Ä‘á»‹nh dáº¡ng tá»« Task!A2 sang Project!A2."
                                        }
                                    },
                                    new()
                                    {
                                        ConditionId = "P22-T1-C04",
                                        Score = 6m,
                                        SourceFile = "xl/worksheets/sheet2.xml",
                                        ExpectedVariants = new List<XmlExpectedVariant>
                                        {
                                            new()
                                            {
                                                ExpectedValues = new List<string>
                                                {
                                                    "<sheetData>"
                                                }
                                            }
                                        },
                                        CompareMode = XmlGradingCompareModes.ExactStringContains,
                                        MatchPolicy = XmlGradingMatchPolicies.All,
                                        Feedback = new ConditionFeedback
                                        {
                                            SuccessDetail = "Worksheet Project cÃ³ dá»¯ liá»‡u XML Ä‘á»ƒ kiá»ƒm tra Ä‘á»‹nh dáº¡ng.",
                                            ErrorMessage = "KhÃ´ng tÃ¬m tháº¥y worksheet XML cáº§n kiá»ƒm tra.",
                                            FixAction = "Kiá»ƒm tra láº¡i sheet Project vÃ  lÆ°u workbook trÆ°á»›c khi cháº¥m láº¡i."
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };
        }
    }
}

