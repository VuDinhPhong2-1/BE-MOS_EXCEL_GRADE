using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
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
        private const string CommonOfficeNamespaceDeclarations =
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
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
                throw new InvalidOperationException($"Project {project.ProjectCode} đã tồn tại trong ruleset.");
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
                throw new InvalidOperationException($"Task {task.TaskId} đã tồn tại trong project {project.ProjectCode}.");
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
                throw new InvalidOperationException($"Condition {condition.ConditionId} đã tồn tại trong task {task.TaskId}.");
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
                ?? throw new InvalidOperationException($"Không tìm thấy XML grading ruleset active cho {subject}/{projectCode}.");

            var rulesetMs = phaseStopwatch.ElapsedMilliseconds;

            var projectRule = ruleSet.Projects.FirstOrDefault(project =>
                string.Equals(project.ProjectCode, NormalizeKey(projectCode), StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Không tìm thấy project rule cho {projectCode}.");

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

                // ===== SPECIAL CONDITION (cấp Task) =====
                // Hoạt động như một "gate" + có điểm riêng: nếu Task có specialCondition
                // mà nó FAIL, toàn bộ Task = 0 điểm bất kể conditions XML thường có đúng
                // hay không. Nếu PASS, cộng thêm SpecialCondition.Score vào điểm Task
                // (độc lập với điểm các Conditions XML, không còn "ăn trọn" MaxScore).
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

                        // Cộng điểm riêng của specialCondition do người tạo ruleset cấu hình.
                        // Cho phép Task chỉ dùng specialCondition (0 condition XML) hoặc
                        // kết hợp cả hai, miễn tổng = task.maxScore (được validate ở ValidateRuleSet).
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
                // Special condition FAIL -> zero toàn bộ Task, kể cả khi có conditions XML đã đạt điểm.
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

                // Nếu FE gửi type rỗng (người dùng chọn "Không sử dụng"), coi như không có special condition.
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

                        // MỚI: dHash luôn là hex 16 ký tự cố định, chỉ cần trim + lowercase,
                        // không cần NormalizeHash (hàm đó dành riêng cho SHA-256).
                        task.SpecialCondition.Config.PerceptualHash = string.IsNullOrWhiteSpace(task.SpecialCondition.Config.PerceptualHash)
                            ? null
                            : task.SpecialCondition.Config.PerceptualHash.Trim().ToLowerInvariant();
                    }

                    // ===== Normalize cho type = insertedImage =====
                    // Tách riêng khối config này với pictureBullet ở trên vì 2 loại
                    // special condition dùng 2 property config khác nhau
                    // (Config vs ImageInsertConfig), không lồng chung 1 object.
                    if (task.SpecialCondition.ImageInsertConfig != null)
                    {
                        var imageInsertConfig = task.SpecialCondition.ImageInsertConfig;

                        imageInsertConfig.AssetId = string.IsNullOrWhiteSpace(imageInsertConfig.AssetId)
                            ? null
                            : imageInsertConfig.AssetId.Trim();

                        imageInsertConfig.ImageHash = string.IsNullOrWhiteSpace(imageInsertConfig.ImageHash)
                            ? null
                            : ImageHashUtility.NormalizeHash(imageInsertConfig.ImageHash);

                        // MỚI
                        imageInsertConfig.PerceptualHash = string.IsNullOrWhiteSpace(imageInsertConfig.PerceptualHash)
                            ? null
                            : imageInsertConfig.PerceptualHash.Trim().ToLowerInvariant();

                        // wrapType là optional: rỗng nghĩa là không cần kiểm tra
                        // chế độ ngắt dòng, chỉ kiểm tra đúng ảnh. Trim + để null
                        // nếu rỗng để tránh lưu chuỗi khoảng trắng vào DB.
                        imageInsertConfig.WrapType = string.IsNullOrWhiteSpace(imageInsertConfig.WrapType)
                            ? null
                            : imageInsertConfig.WrapType.Trim();
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
                throw new InvalidOperationException("Subject là bắt buộc.");
            }

            if (string.IsNullOrWhiteSpace(ruleSet.Version))
            {
                throw new InvalidOperationException("Version là bắt buộc.");
            }
        }

        private static void ValidateProjectShell(ProjectXmlRule project)
        {
            if (string.IsNullOrWhiteSpace(project.ProjectCode))
            {
                throw new InvalidOperationException("ProjectCode là bắt buộc.");
            }

            if (project.MaxScore <= 0)
            {
                throw new InvalidOperationException("Project maxScore phải lớn hơn 0.");
            }
        }

        private static void ValidateTaskShell(TaskXmlRule task)
        {
            if (string.IsNullOrWhiteSpace(task.TaskId))
            {
                throw new InvalidOperationException("TaskId là bắt buộc.");
            }

            if (task.MaxScore <= 0)
            {
                throw new InvalidOperationException("Task maxScore phải lớn hơn 0.");
            }

            if (task.SpecialCondition != null)
            {
                if (!SpecialConditionTypes.Supported.Contains(task.SpecialCondition.Type))
                {
                    throw new InvalidOperationException($"specialCondition.type không được hỗ trợ: {task.SpecialCondition.Type}.");
                }

                if (task.SpecialCondition.Score <= 0)
                {
                    throw new InvalidOperationException("specialCondition.score phải lớn hơn 0.");
                }
            }
        }

        private static void ValidateConditionShell(XmlGradingCondition condition)
        {
            if (string.IsNullOrWhiteSpace(condition.ConditionId))
            {
                throw new InvalidOperationException("ConditionId là bắt buộc.");
            }

            if (condition.Score <= 0)
            {
                throw new InvalidOperationException("Condition score phải lớn hơn 0.");
            }

            if (!IsSafeSourceFile(condition.SourceFile))
            {
                throw new InvalidOperationException("sourceFile phải là đường dẫn XML an toàn trong Office package.");
            }

            if (condition.ExpectedVariants == null || condition.ExpectedVariants.Count == 0)
            {
                throw new InvalidOperationException("expectedVariants là bắt buộc.");
            }

            if (condition.ExpectedVariants.Any(variant => variant == null || variant.ExpectedValues.Count == 0))
            {
                throw new InvalidOperationException("Mỗi expectedVariant phải có ít nhất 1 expectedValues.");
            }

            if (!XmlGradingCompareModes.Supported.Contains(condition.CompareMode))
            {
                throw new InvalidOperationException($"compareMode không hỗ trợ: {condition.CompareMode}.");
            }

            if (!XmlGradingMatchPolicies.Supported.Contains(condition.MatchPolicy))
            {
                throw new InvalidOperationException($"matchPolicy không hỗ trợ: {condition.MatchPolicy}.");
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
                    || string.Equals(specialConditionType, SpecialConditionTypes.DocumentStyleSet, StringComparison.OrdinalIgnoreCase),
                "excel" => false,
                "ppt" => false,
                "powerpoint" => false,
                _ => false
            };
        }

        private sealed class RequiredOfficeParts
        {
            public HashSet<string> XmlParts { get; } = new(StringComparer.OrdinalIgnoreCase);
            public bool ReadRelatedImages { get; set; }
        }

        private sealed class OfficePackage
        {
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

            if (requiredParts == null || requiredParts.XmlParts.Count == 0)
            {
                foreach (var entry in archive.Entries)
                {
                    var normalizedPath = NormalizeSourceFile(entry.FullName);

                    if (string.IsNullOrWhiteSpace(normalizedPath))
                    {
                        continue;
                    }

                    if (normalizedPath.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        || normalizedPath.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
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
                    AddXmlPart("word/document.xml", "word/document.xml");
                    AddXmlPart("word/_rels/document.xml.rels", "word/_rels/document.xml.rels");
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
                    result.Feedback.ErrorMessage = $"Không tìm thấy XML part {result.SourceFile} trong file học sinh.";
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
        /// Kết quả nội bộ khi đánh giá 1 Special Condition (không phơi ra ngoài API,
        /// chỉ dùng để build message cho taskResult.Details/Errors).
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

            // MỚI
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

            return new SpecialConditionEvalOutcome
            {
                IsPassed = false,
                Message = $"Special condition không được hỗ trợ: {specialCondition.Type}."
            };
        }

        /// <summary>
        /// So sánh ảnh thực tế với ảnh chuẩn (expectedHash/expectedPerceptualHash).
        /// Ưu tiên perceptual hash (chịu được Word nén lại JPEG khi save); nếu
        /// ruleset chưa có perceptualHash (dữ liệu cũ), hoặc ảnh không decode
        /// được (vd .wmf/.emf), fallback về so SHA-256 tuyệt đối như trước.
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
                    // Ảnh không decode được bằng ImageSharp (vd .wmf/.emf) -> fallback SHA-256
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
                        .Select(paragraph => BuildParagraphTextSnapshot(paragraph, w).Text)
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

        private static ParagraphTextSnapshot BuildParagraphTextSnapshot(XElement paragraph, XNamespace w)
        {
            var text = new StringBuilder();
            var tabCount = 0;

            foreach (var node in paragraph.Descendants())
            {
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
                return Fail("Chưa cấu hình Inserted Image (config trống).");
            }

            if (string.IsNullOrWhiteSpace(config.ImageHash))
            {
                return Fail("Chưa có imageHash — ảnh chuẩn chưa được upload/tạo hash ở BE.");
            }

            const string documentPart = "word/document.xml";
            const string documentRelsPath = "word/_rels/document.xml.rels";

            if (!package.TryGetXmlDocument(documentPart, out var documentDocument, out var documentError))
            {
                return Fail($"Không tìm thấy {documentPart} trong file học sinh.");
            }

            if (!package.TryGetRelationships(documentRelsPath, out var relationships, out var relsError))
            {
                return Fail("Không tìm thấy word/_rels/document.xml.rels.");
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

                var drawings = documentDocument.Descendants(w + "drawing").ToList();

                if (drawings.Count == 0)
                {
                    return Fail("Không tìm thấy hình ảnh (w:drawing) nào trong tài liệu.");
                }

                string? lastMismatchInfo = null;

                foreach (var drawing in drawings)
                {
                    var relationshipId = drawing.Descendants(a + "blip").FirstOrDefault()?.Attribute(r + "embed")?.Value;

                    if (string.IsNullOrWhiteSpace(relationshipId))
                    {
                        continue;
                    }

                    if (!relationships.TryGetValue(relationshipId, out var target) || string.IsNullOrWhiteSpace(target))
                    {
                        lastMismatchInfo = $"Relationship {relationshipId} không có Target hợp lệ.";
                        continue;
                    }

                    var imagePath = ResolveRelationshipTarget(documentPart, target);

                    if (!package.BinaryParts.TryGetValue(imagePath, out var imageBytes))
                    {
                        lastMismatchInfo = $"Không đọc được ảnh {imagePath} trong file học sinh.";
                        continue;
                    }

                    if (!IsImageMatch(package, imagePath, imageBytes, expectedHash, expectedPerceptualHash))
                    {
                        lastMismatchInfo = $"Tìm thấy ảnh {imagePath} nhưng không đúng nội dung yêu cầu.";
                        continue;
                    }

                    if (expectedWrap == null)
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = "Đã chèn đúng hình ảnh yêu cầu."
                        };
                    }

                    var actualWrap = DetectWrapType(drawing, wp);

                    if (string.Equals(actualWrap, expectedWrap, StringComparison.OrdinalIgnoreCase))
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = $"Đã chèn đúng hình ảnh yêu cầu với chế độ ngắt dòng '{expectedWrap}'."
                        };
                    }

                    lastMismatchInfo = $"Đã tìm thấy đúng ảnh yêu cầu nhưng chế độ ngắt dòng đang là '{actualWrap}' thay vì '{expectedWrap}'.";
                }

                return Fail(lastMismatchInfo ?? "Không tìm thấy ảnh nào khớp với yêu cầu trong tài liệu.");
            }
            catch (XmlException ex)
            {
                return Fail($"Không thể phân tích XML: {ex.Message}");
            }
        }

        /// <summary>
        /// wp:inline = "inline". wp:anchor có 1 trong các phần tử wrap con:
        /// wrapSquare/wrapTight/wrapThrough/wrapTopAndBottom/wrapNone
        /// (behind/inFront phân biệt bằng attribute behindDoc trên wp:anchor).
        /// </summary>
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
                return Fail("Chưa cấu hình Picture Bullet (config trống).");
            }

            if (string.IsNullOrWhiteSpace(config.ImageHash))
            {
                return Fail("Chưa có imageHash — ảnh bullet chuẩn chưa được upload/tạo hash ở BE.");
            }

            const string documentPart = "word/document.xml";
            const string numberingPath = "word/numbering.xml";
            const string numberingRelsPath = "word/_rels/numbering.xml.rels";

            if (!package.TryGetXmlDocument(documentPart, out var documentDocument, out var documentError))
            {
                return Fail($"Không tìm thấy {documentPart} trong file học sinh.");
            }

            if (!package.TryGetXmlDocument(numberingPath, out var numberingDocument, out var numberingError))
            {
                return Fail("File học sinh không có word/numbering.xml.");
            }

            if (!package.TryGetRelationships(numberingRelsPath, out var relationships, out var relsError))
            {
                return Fail("Không tìm thấy word/_rels/numbering.xml.rels.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            try
            {
                var expectedHash = ImageHashUtility.NormalizeHash(config.ImageHash);
                var expectedPerceptualHash = config.PerceptualHash;

                // Duyệt toàn bộ paragraph trong document.xml, tìm bất kỳ paragraph nào
                // dùng picture bullet ở đúng Level (nếu có chỉ định) mà ảnh khớp expectedHash.
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
                        // Level này không dùng picture bullet -> không tính là "đã kiểm tra bullet"
                        continue;
                    }

                    checkedAnyBulletParagraph = true;

                    // lvlPicBulletId -> numPicBullet -> r:id ảnh
                    if (!numPicBulletById.TryGetValue(lvlPicBulletId.ToString(), out var numPicBulletElement))
                    {
                        lastMismatchInfo = $"numPicBulletId={lvlPicBulletId} không tồn tại trong numbering.xml.";
                        continue;
                    }

                    var relationshipId =
                        numPicBulletElement.Descendants(v + "imagedata").FirstOrDefault()?.Attribute(r + "id")?.Value
                        ?? numPicBulletElement.Descendants(a + "blip").FirstOrDefault()?.Attribute(r + "embed")?.Value;

                    if (string.IsNullOrWhiteSpace(relationshipId))
                    {
                        lastMismatchInfo = $"numPicBulletId={lvlPicBulletId} không có tham chiếu ảnh.";
                        continue;
                    }

                    if (!relationships.TryGetValue(relationshipId, out var target) || string.IsNullOrWhiteSpace(target))
                    {
                        lastMismatchInfo = $"Relationship {relationshipId} không có Target hợp lệ.";
                        continue;
                    }

                    var imagePath = ResolveRelationshipTarget(numberingPath, target);

                    if (!package.BinaryParts.TryGetValue(imagePath, out var imageBytes))
                    {
                        lastMismatchInfo = $"Không đọc được ảnh {imagePath} trong file học sinh.";
                        continue;
                    }

                    if (IsImageMatch(package, imagePath, imageBytes, expectedHash, expectedPerceptualHash))
                    {
                        return new SpecialConditionEvalOutcome
                        {
                            IsPassed = true,
                            Message = "Picture bullet đúng hình ảnh yêu cầu."
                        };
                    }

                    lastMismatchInfo = $"Tìm thấy picture bullet (numId={numId}, level={ilvl}) nhưng ảnh không khớp.";
                }

                if (!checkedAnyBulletParagraph)
                {
                    return Fail(config.Level.HasValue
                        ? $"Không tìm thấy paragraph nào dùng picture bullet ở level {config.Level.Value}."
                        : "Không tìm thấy paragraph nào dùng picture bullet trong tài liệu.");
                }

                return Fail(lastMismatchInfo ?? "Picture bullet không đúng hình ảnh yêu cầu.");
            }
            catch (XmlException ex)
            {
                return Fail($"Không thể phân tích XML: {ex.Message}");
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

            // Không nên âm thầm coi mode lạ là normalized
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

            // Chuẩn hóa search space 1 lần duy nhất (không đổi thứ tự ký tự nên cursor vẫn hợp lệ)
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

                // Tìm bắt đầu từ cursor hiện tại -> đảm bảo đúng thứ tự và liền mạch,
                // không cho phép match ở một occurrence đứng trước fragment trước đó.
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
                // Nếu không match, giữ nguyên cursor -> các fragment sau vẫn được thử,
                // nhưng IsPassed cuối cùng sẽ = false vì có ít nhất 1 fragment missing.
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

            // Chuẩn hóa khoảng trắng
            text = Regex.Replace(
                text,
                @"\s+",
                " ");

            // Chuẩn hóa khoảng trắng quanh dấu =
            text = Regex.Replace(
                text,
                @"\s*=\s*",
                "=");

            // Chuẩn hóa khoảng trắng trước />
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

        private static int? ParseIntAttribute(XElement element, string localName)
        {
            var value = element.Attributes().FirstOrDefault(attribute =>
                string.Equals(attribute.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase))?.Value;

            return int.TryParse(value, out var parsed) ? parsed : null;
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
                    matches.All(match => match.IsMatched), // thứ tự đã được đảm bảo bởi cursor khi tìm kiếm

                _ => matches.All(match => match.IsMatched)
            };
        }

        private static XmlRuleValidationResult ValidateRuleSet(GradingRuleSet ruleSet)
        {
            var result = new XmlRuleValidationResult();

            if (string.IsNullOrWhiteSpace(ruleSet.Subject))
            {
                result.Errors.Add("subject không được rỗng.");
            }

            if (ruleSet.Projects.Count == 0)
            {
                result.Errors.Add("projects phải có ít nhất 1 project.");
            }

            foreach (var project in ruleSet.Projects)
            {
                var projectPrefix = string.IsNullOrWhiteSpace(project.ProjectCode) ? "project" : project.ProjectCode;
                if (string.IsNullOrWhiteSpace(project.ProjectCode))
                {
                    result.Errors.Add("project.projectCode không được rỗng.");
                }

                if (project.MaxScore <= 0)
                {
                    result.Errors.Add($"{projectPrefix}.maxScore phải lớn hơn 0.");
                }

                foreach (var task in project.Tasks)
                {
                    var taskPrefix = string.IsNullOrWhiteSpace(task.TaskId) ? $"{projectPrefix}.task" : $"{projectPrefix}.{task.TaskId}";

                    if (string.IsNullOrWhiteSpace(task.TaskId))
                    {
                        result.Errors.Add($"{projectPrefix}.taskId không được rỗng.");
                    }

                    if (string.IsNullOrWhiteSpace(task.TaskName))
                    {
                        result.Errors.Add($"{taskPrefix}.taskName không được rỗng.");
                    }

                    if (task.MaxScore <= 0)
                    {
                        result.Errors.Add($"{taskPrefix}.maxScore phải lớn hơn 0.");
                    }

                    var hasSpecialCondition = task.SpecialCondition != null
                        && !string.IsNullOrWhiteSpace(task.SpecialCondition.Type);

                    // Task hợp lệ khi có ít nhất 1 Condition XML HOẶC có specialCondition
                    // (không còn bắt buộc phải có Condition XML nếu đã dùng specialCondition).
                    if (task.Conditions.Count == 0 && !hasSpecialCondition)
                    {
                        result.Errors.Add($"{taskPrefix}.conditions phải có ít nhất 1 condition, hoặc phải có specialCondition.");
                    }
                    else
                    {
                        // Tổng điểm = tổng Conditions XML + điểm riêng của specialCondition (nếu có).
                        // Cho phép Task chỉ dùng specialCondition (0 Condition XML) hoặc kết hợp cả hai.
                        var totalConditionScore = task.Conditions.Sum(condition => condition.Score)
                            + (hasSpecialCondition ? task.SpecialCondition!.Score : 0m);

                        if (totalConditionScore != task.MaxScore)
                        {
                            var scoreBreakdown = hasSpecialCondition
                                ? $"conditions + specialCondition = {totalConditionScore}"
                                : $"conditions = {totalConditionScore}";

                            result.Errors.Add($"{taskPrefix}: tổng score ({scoreBreakdown}) phải bằng task.maxScore ({task.MaxScore}).");
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
                result.Errors.Add($"{taskPrefix}.conditionId không được rỗng.");
            }

            if (condition.Score <= 0)
            {
                result.Errors.Add($"{conditionPrefix}.score phải lớn hơn 0.");
            }

            if (string.IsNullOrWhiteSpace(condition.SourceFile))
            {
                result.Errors.Add($"{conditionPrefix}.sourceFile không được rỗng.");
            }
            else if (!IsSafeSourceFile(condition.SourceFile))
            {
                result.Errors.Add($"{conditionPrefix}.sourceFile không hợp lệ hoặc có path traversal.");
            }

            if (condition.ExpectedVariants == null || condition.ExpectedVariants.Count == 0)
            {
                result.Errors.Add($"{conditionPrefix}.expectedVariants phải có ít nhất 1 variant.");
            }
            else
            {
                for (var variantIndex = 0; variantIndex < condition.ExpectedVariants.Count; variantIndex++)
                {
                    var variant = condition.ExpectedVariants[variantIndex];
                    if (variant == null || variant.ExpectedValues.Count == 0 || variant.ExpectedValues.Any(string.IsNullOrWhiteSpace))
                    {
                        result.Errors.Add($"{conditionPrefix}.expectedVariants[{variantIndex}].expectedValues phải là array string không rỗng.");
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(condition.CompareMode))
            {
                condition.CompareMode = XmlGradingCompareModes.XmlContainsNormalized;
            }

            if (!XmlGradingCompareModes.Supported.Contains(condition.CompareMode))
            {
                result.Errors.Add($"{conditionPrefix}.compareMode không được hỗ trợ: {condition.CompareMode}.");
            }

            if (string.IsNullOrWhiteSpace(condition.MatchPolicy))
            {
                condition.MatchPolicy = XmlGradingMatchPolicies.All;
            }

            if (!XmlGradingMatchPolicies.Supported.Contains(condition.MatchPolicy))
            {
                result.Errors.Add($"{conditionPrefix}.matchPolicy không được hỗ trợ: {condition.MatchPolicy}.");
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
                result.Errors.Add($"{conditionPrefix}.xmlEquivalentWholeFile chỉ hỗ trợ đúng 1 expectedValue trong mỗi expectedVariant.");
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
                result.Errors.Add($"{taskPrefix}.specialCondition.type không được rỗng.");
                return;
            }

            if (!SpecialConditionTypes.Supported.Contains(specialCondition.Type))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.type không được hỗ trợ: {specialCondition.Type}.");
                return;
            }

            if (!IsSpecialConditionSupportedForSubject(specialCondition.Type, subject))
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.type {specialCondition.Type} khong ho tro cho subject {NormalizeKey(subject)}.");
                return;
            }

            if (specialCondition.Score <= 0)
            {
                result.Errors.Add($"{taskPrefix}.specialCondition.score phải lớn hơn 0.");
            }

            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureBullet, StringComparison.OrdinalIgnoreCase))
            {
                var config = specialCondition.Config;

                if (config == null)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.config không được null.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(config.ImageHash))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.config.imageHash không được rỗng (ảnh bullet chuẩn chưa được upload/tạo hash).");
                }

                if (config.Level.HasValue && config.Level.Value < 0)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.config.level phải >= 0.");
                }
            }

            // MỚI
            if (string.Equals(specialCondition.Type, SpecialConditionTypes.InsertedImage, StringComparison.OrdinalIgnoreCase))
            {
                var config = specialCondition.ImageInsertConfig;

                if (config == null)
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.imageInsertConfig không được null.");
                    return;
                }

                if (string.IsNullOrWhiteSpace(config.ImageHash))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.imageInsertConfig.imageHash không được rỗng (ảnh chuẩn chưa được upload/tạo hash).");
                }

                if (!string.IsNullOrWhiteSpace(config.WrapType) && !ImageWrapTypes.Supported.Contains(config.WrapType))
                {
                    result.Errors.Add($"{taskPrefix}.specialCondition.imageInsertConfig.wrapType không được hỗ trợ: {config.WrapType}.");
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

            // Phải là đường dẫn tương đối trong Office ZIP package
            if (Path.IsPathRooted(normalized))
                return false;

            // Không được bắt đầu bằng /
            if (normalized.StartsWith("/", StringComparison.Ordinal))
                return false;

            // Không cho phép path traversal
            var segments = normalized.Split(
                '/',
                StringSplitOptions.RemoveEmptyEntries);

            if (segments.Any(segment =>
                segment == ".."))
            {
                return false;
            }

            // Không cho phép segment rỗng bất thường
            if (segments.Any(segment =>
                segment == "."))
            {
                return false;
            }

            // Chỉ cho phép XML hoặc RELS
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

            // Loại bỏ ./ ở đầu
            while (normalized.StartsWith("./", StringComparison.Ordinal))
            {
                normalized = normalized[2..];
            }

            // Không cho phép / ở đầu
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

            // MaxScore của bài = tổng MaxScore do người tạo
            // cấu hình cho từng Task.
            result.MaxScore = gradableTasks.Sum(task => task.MaxScore);

            // Điểm thực tế = tổng điểm các Task đạt được.
            result.TotalScore = Math.Round(
                gradableTasks.Sum(task => task.Score),
                2,
                MidpointRounding.AwayFromZero
            );

            // Không cho vượt quá điểm tối đa.
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
                                TaskName = "Sao chép định dạng từ tiêu đề và phụ đề của trang tính Task sang Project.",
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
                                            SuccessDetail = "Đã xác nhận XML tiêu đề/phụ đề nguồn tại Task!A1:A2.",
                                            ErrorMessage = "Không tìm thấy đầy đủ XML tiêu đề/phụ đề nguồn tại Task!A1:A2.",
                                            FixAction = "Kiểm tra lại nội dung tiêu đề/phụ đề trong sheet Task."
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
                                            SuccessDetail = "Project!A1 có XML định dạng đúng.",
                                            ErrorMessage = "Project!A1 chưa có XML định dạng đúng.",
                                            FixAction = "Dùng Format Painter sao chép định dạng từ Task!A1 sang Project!A1."
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
                                            SuccessDetail = "Project!A2 có XML định dạng đúng.",
                                            ErrorMessage = "Project!A2 chưa có XML định dạng đúng.",
                                            FixAction = "Dùng Format Painter sao chép định dạng từ Task!A2 sang Project!A2."
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
                                            SuccessDetail = "Worksheet Project có dữ liệu XML để kiểm tra định dạng.",
                                            ErrorMessage = "Không tìm thấy worksheet XML cần kiểm tra.",
                                            FixAction = "Kiểm tra lại sheet Project và lưu workbook trước khi chấm lại."
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
