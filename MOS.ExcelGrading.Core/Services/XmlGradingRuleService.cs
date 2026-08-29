using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;
namespace MOS.ExcelGrading.Core.Services
{
    public class XmlGradingRuleService : IXmlGradingRuleService
    {
        private const decimal StandardProjectMaxScore = 125m;
        private readonly IMongoCollection<GradingRuleSet> _ruleSets;

        public XmlGradingRuleService(IMongoDatabase database)
        {
            _ruleSets = database.GetCollection<GradingRuleSet>("grading_rule_sets");
        }

        public async Task<List<GradingRuleSet>> GetRuleSetsAsync(string? subject = null, bool? isActive = null)
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

            var filter = filters.Count == 0
                ? Builders<GradingRuleSet>.Filter.Empty
                : Builders<GradingRuleSet>.Filter.And(filters);

            return await _ruleSets.Find(filter).ToListAsync();
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

            return await _ruleSets.Find(filter).FirstOrDefaultAsync();
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
            var ruleSet = await GetActiveRuleSetAsync(subject, projectCode)
                ?? throw new InvalidOperationException($"Không tìm thấy XML grading ruleset active cho {subject}/{projectCode}.");

            var projectRule = ruleSet.Projects.FirstOrDefault(project =>
                string.Equals(project.ProjectCode, NormalizeKey(projectCode), StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Không tìm thấy project rule cho {projectCode}.");

            var package = ReadOfficePackage(studentFile);

            var result = new GradingResult
            {
                ProjectId = projectRule.ProjectCode,
                ProjectName = string.IsNullOrWhiteSpace(projectRule.ProjectName) ? projectRule.ProjectCode : projectRule.ProjectName,
                MaxScore = StandardProjectMaxScore
            };

            foreach (var taskRule in projectRule.Tasks)
            {
                var taskResult = new TaskResult
                {
                    TaskId = taskRule.TaskId,
                    TaskName = taskRule.TaskName,
                    MaxScore = taskRule.MaxScore
                };

                foreach (var condition in taskRule.Conditions)
                {
                    var conditionResult = EvaluateCondition(condition, package);
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
                        if (!string.IsNullOrWhiteSpace(conditionResult.Feedback.ErrorMessage))
                        {
                            taskResult.Errors.Add(conditionResult.Feedback.ErrorMessage.Trim());
                        }

                        if (!string.IsNullOrWhiteSpace(conditionResult.Feedback.FixAction))
                        {
                            taskResult.FixActions.Add(conditionResult.Feedback.FixAction.Trim());
                        }
                    }

                    taskResult.Details.Add(
                        $"XML condition {conditionResult.ConditionId}: sourceFile={conditionResult.SourceFile}, compareMode={conditionResult.CompareMode}, matchPolicy={conditionResult.MatchPolicy}, matched={conditionResult.MatchedExpectedValues.Count}, missing={conditionResult.MissingExpectedValues.Count}.");

                    if (!conditionResult.IsPassed && condition.StopTaskIfFailed)
                    {
                        break;
                    }
                }

                if (taskResult.Score > taskResult.MaxScore)
                {
                    taskResult.Score = taskResult.MaxScore;
                }

                result.TaskResults.Add(taskResult);
            }

            ApplyProjectScoringModel(result);
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
        }

        private static void NormalizeConditionForPersistence(XmlGradingCondition condition)
        {
            condition.ConditionId = condition.ConditionId?.Trim() ?? string.Empty;
            condition.SourceFile = NormalizeSourceFile(condition.SourceFile);
            condition.ExpectedValues = condition.ExpectedValues?
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .ToList() ?? new List<string>();
            condition.CompareMode = string.IsNullOrWhiteSpace(condition.CompareMode)
                ? XmlGradingCompareModes.XmlContainsNormalized
                : condition.CompareMode.Trim();
            condition.MatchPolicy = string.IsNullOrWhiteSpace(condition.MatchPolicy)
                ? XmlGradingMatchPolicies.All
                : condition.MatchPolicy.Trim();
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

    var hasSpecialCondition = condition.SpecialCondition != null
        && condition.SpecialCondition.Type != SpecialConditionType.None;

    // sourceFile: nếu có special condition mà không có expectedValues,
    // sourceFile là tùy chọn (special condition tự quản lý documentPart riêng)
    var hasNormalXmlCheck = condition.ExpectedValues.Count > 0;

    if (hasNormalXmlCheck || !hasSpecialCondition)
    {
        if (string.IsNullOrWhiteSpace(condition.SourceFile))
        {
            result.Errors.Add($"{conditionPrefix}.sourceFile không được rỗng.");
        }
        else if (!IsSafeSourceFile(condition.SourceFile))
        {
            result.Errors.Add($"{conditionPrefix}.sourceFile không hợp lệ hoặc có path traversal.");
        }

        if (condition.ExpectedValues.Count == 0 || condition.ExpectedValues.Any(string.IsNullOrWhiteSpace))
        {
            // Chỉ bắt buộc expectedValues khi KHÔNG có special condition thay thế
            if (!hasSpecialCondition)
            {
                result.Errors.Add($"{conditionPrefix}.expectedValue phải là string không rỗng hoặc array string không rỗng.");
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

    if (string.Equals(condition.CompareMode, XmlGradingCompareModes.XmlEquivalentWholeFile, StringComparison.OrdinalIgnoreCase) &&
        condition.ExpectedValues.Count != 1)
    {
        result.Errors.Add($"{conditionPrefix}.xmlEquivalentWholeFile chỉ hỗ trợ đúng 1 expectedValue.");
    }

    ValidateSpecialCondition(condition.SpecialCondition, conditionPrefix, result);

    if (!hasSpecialCondition && condition.ExpectedValues.Count == 0)
    {
        result.Errors.Add($"{conditionPrefix} phải có expectedValues hoặc specialCondition.");
    }
}


        private sealed class OfficePackage
        {
            public Dictionary<string, string> XmlParts { get; } =
                new(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, byte[]> BinaryParts { get; } =
                new(StringComparer.OrdinalIgnoreCase);
        }
        private static OfficePackage ReadOfficePackage(Stream studentFile)
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

    foreach (var entry in archive.Entries)
    {
        var normalizedPath = NormalizeSourceFile(entry.FullName);

        if (string.IsNullOrWhiteSpace(normalizedPath))
        {
            continue;
        }

        using var entryStream = entry.Open();

        if (normalizedPath.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || normalizedPath.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
        {
            using var reader = new StreamReader(
                entryStream,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true);

            package.XmlParts[normalizedPath] = reader.ReadToEnd();
        }
        else if (IsSupportedImage(normalizedPath))
        {
            using var memoryStream = new MemoryStream();
            entryStream.CopyTo(memoryStream);
            package.BinaryParts[normalizedPath] = memoryStream.ToArray();
        }
    }

    return package;
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
            OfficePackage package)
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

            // ===== SPECIAL CONDITION =====
            if (condition.SpecialCondition != null
                && condition.SpecialCondition.Type != SpecialConditionType.None)
            {
                var specialResult = EvaluateSpecialCondition(condition.SpecialCondition, package);
                result.SpecialConditionResult = specialResult;

                if (!specialResult.IsPassed)
                {
                    result.IsPassed = false;
                    result.ScoreAwarded = 0m;

                    if (string.IsNullOrWhiteSpace(result.Feedback.ErrorMessage))
                    {
                        result.Feedback.ErrorMessage = specialResult.Message;
                    }

                    return result;
                }

                // Không có normal XML check đi kèm -> Special Condition PASS là đủ
                if (condition.ExpectedValues.Count == 0)
                {
                    result.IsPassed = true;
                    result.ScoreAwarded = condition.Score;

                    if (string.IsNullOrWhiteSpace(result.Feedback.SuccessDetail))
                    {
                        result.Feedback.SuccessDetail = specialResult.Message;
                    }

                    return result;
                }
            }

            // ===== NORMAL XML CONDITION =====
            if (!package.XmlParts.TryGetValue(result.SourceFile, out var actualXml))
            {
                result.MissingExpectedValues.AddRange(condition.ExpectedValues);
                if (string.IsNullOrWhiteSpace(result.Feedback.ErrorMessage))
                {
                    result.Feedback.ErrorMessage = $"Không tìm thấy XML part {result.SourceFile} trong file học sinh.";
                }

                return result;
            }

            var matches = condition.ExpectedValues
                .Select(expected => MatchExpected(actualXml, expected, compareMode))
                .ToList();

            var isPassed = ApplyMatchPolicy(matches, matchPolicy);
            result.IsPassed = isPassed;
            result.ScoreAwarded = isPassed ? condition.Score : 0m;
            result.MatchedExpectedValues = matches.Where(m => m.IsMatched).Select(m => m.ExpectedValue).ToList();
            result.MissingExpectedValues = matches.Where(m => !m.IsMatched).Select(m => m.ExpectedValue).ToList();

            return result;
        }

        private static SpecialConditionEvaluationResult EvaluateSpecialCondition(

// ===== NORMAL XML CONDITION ===== (giữ nguyên như cũ)

    private static SpecialConditionEvaluationResult EvaluateSpecialCondition(
    SpecialCondition specialCondition,
    OfficePackage package)
{
    return specialCondition.Type switch
    {
        SpecialConditionType.PictureBullet =>
            EvaluatePictureBullet(specialCondition.PictureBullet, package),

        _ => new SpecialConditionEvaluationResult
        {
            IsPassed = false,
            Type = specialCondition.Type.ToString(),
            Message = $"Special condition không được hỗ trợ: {specialCondition.Type}."
        }
    };
}

private static SpecialConditionEvaluationResult EvaluatePictureBullet(
    PictureBulletConfig? config,
    OfficePackage package)
{
    static SpecialConditionEvaluationResult Fail(string message) => new()
    {
        IsPassed = false,
        Type = "PictureBullet",
        Message = message
    };

    if (config == null)
    {
        return Fail("PictureBulletConfig không được cấu hình.");
    }

    if (string.IsNullOrWhiteSpace(config.ExpectedImageSha256))
    {
        return Fail("Chưa cấu hình ExpectedImageSha256.");
    }

    var documentPart = NormalizeSourceFile(
        string.IsNullOrWhiteSpace(config.DocumentPart)
            ? "word/document.xml"
            : config.DocumentPart);

    if (!package.XmlParts.TryGetValue(documentPart, out var documentXml))
    {
        return Fail($"Không tìm thấy {documentPart} trong file học sinh.");
    }

    const string numberingPath = "word/numbering.xml";
    const string numberingRelsPath = "word/_rels/numbering.xml.rels";

    if (!package.XmlParts.TryGetValue(numberingPath, out var numberingXml))
    {
        return Fail("File học sinh không có word/numbering.xml.");
    }

    if (!package.XmlParts.TryGetValue(numberingRelsPath, out var relsXml))
    {
        return Fail("Không tìm thấy word/_rels/numbering.xml.rels.");
    }

    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    XNamespace v = "urn:schemas-microsoft-com:vml";
    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    try
    {
        var documentDocument = XDocument.Parse(documentXml);
        var numberingDocument = XDocument.Parse(numberingXml);
        var relsDocument = XDocument.Parse(relsXml);

        // Bước 1: lấy paragraph theo index (đếm toàn bộ w:p, kể cả không có numPr)
        var allParagraphs = documentDocument.Descendants(w + "p").ToList();

        List<XElement> paragraphsToCheck;

        if (config.ParagraphIndex.HasValue)
        {
            if (config.ParagraphIndex.Value < 0
                || config.ParagraphIndex.Value >= allParagraphs.Count)
            {
                return Fail(
                    $"paragraphIndex {config.ParagraphIndex.Value} vượt ngoài phạm vi document ({allParagraphs.Count} paragraph).");
            }

            paragraphsToCheck = new List<XElement> { allParagraphs[config.ParagraphIndex.Value] };
        }
        else
        {
            paragraphsToCheck = allParagraphs;
        }

        var expectedHash = NormalizeHash(config.ExpectedImageSha256);
        SpecialConditionEvaluationResult? lastAttempt = null;

        foreach (var paragraph in paragraphsToCheck)
        {
            var numPr = paragraph.Element(w + "pPr")?.Element(w + "numPr");
            if (numPr == null)
            {
                continue; // paragraph không có numbering
            }

            var numIdVal = numPr.Element(w + "numId")?.Attribute(w + "val")?.Value;
            if (!int.TryParse(numIdVal, out var numId))
            {
                continue;
            }

            if (config.NumId.HasValue && numId != config.NumId.Value)
            {
                continue;
            }

            // ilvl mặc định = 0 nếu không khai báo
            var ilvlVal = numPr.Element(w + "ilvl")?.Attribute(w + "val")?.Value;
            var ilvl = int.TryParse(ilvlVal, out var parsedIlvl) ? parsedIlvl : 0;

            if (config.Level.HasValue && ilvl != config.Level.Value)
            {
                continue;
            }

            // Bước 2: numId -> w:num -> abstractNumId
            var numElement = numberingDocument
                .Descendants(w + "num")
                .FirstOrDefault(e => e.Attribute(w + "numId")?.Value == numId.ToString());

            if (numElement == null)
            {
                lastAttempt = Fail($"Không tìm thấy w:num numId={numId} trong numbering.xml.");
                continue;
            }

            var abstractNumIdVal = numElement.Element(w + "abstractNumId")?.Attribute(w + "val")?.Value;
            if (!int.TryParse(abstractNumIdVal, out var abstractNumId))
            {
                lastAttempt = Fail($"numId={numId} thiếu w:abstractNumId hợp lệ.");
                continue;
            }

            // Bước 3: abstractNumId -> abstractNum -> lvl[ilvl] -> lvlPicBulletId
            var abstractNumElement = numberingDocument
                .Descendants(w + "abstractNum")
                .FirstOrDefault(e => e.Attribute(w + "abstractNumId")?.Value == abstractNumId.ToString());

            if (abstractNumElement == null)
            {
                lastAttempt = Fail($"Không tìm thấy w:abstractNum abstractNumId={abstractNumId}.");
                continue;
            }

            var lvlElement = abstractNumElement
                .Elements(w + "lvl")
                .FirstOrDefault(e => e.Attribute(w + "ilvl")?.Value == ilvl.ToString());

            if (lvlElement == null)
            {
                lastAttempt = Fail($"Không tìm thấy w:lvl ilvl={ilvl} trong abstractNum {abstractNumId}.");
                continue;
            }

            var lvlPicBulletIdVal = lvlElement.Element(w + "lvlPicBulletId")?.Attribute(w + "val")?.Value;
            if (!int.TryParse(lvlPicBulletIdVal, out var lvlPicBulletId))
            {
                lastAttempt = Fail(
                    $"Level {ilvl} của numId={numId} không dùng picture bullet (thiếu w:lvlPicBulletId).");
                continue;
            }

            // Bước 4: lvlPicBulletId -> numPicBullet -> r:id ảnh
            var numPicBulletElement = numberingDocument
                .Descendants(w + "numPicBullet")
                .FirstOrDefault(e => e.Attribute(w + "numPicBulletId")?.Value == lvlPicBulletId.ToString());

            if (numPicBulletElement == null)
            {
                lastAttempt = Fail($"Không tìm thấy w:numPicBullet numPicBulletId={lvlPicBulletId}.");
                continue;
            }

            // Hỗ trợ cả 2 dạng Word hay sinh: VML (v:imagedata) và DrawingML (a:blip)
            var relationshipId =
                numPicBulletElement.Descendants(v + "imagedata").FirstOrDefault()?.Attribute(r + "id")?.Value
                ?? numPicBulletElement.Descendants(a + "blip").FirstOrDefault()?.Attribute(r + "embed")?.Value;

            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                lastAttempt = Fail($"numPicBullet {lvlPicBulletId} không có tham chiếu ảnh (v:imagedata / a:blip).");
                continue;
            }

            // Bước 5: relationship id -> target -> resolve path -> bytes -> sha256
            var relationship = relsDocument
                .Descendants(rel + "Relationship")
                .FirstOrDefault(e => string.Equals(
                    e.Attribute("Id")?.Value, relationshipId, StringComparison.Ordinal));

            if (relationship == null)
            {
                lastAttempt = Fail($"Không tìm thấy relationship {relationshipId} trong numbering.xml.rels.");
                continue;
            }

            var target = relationship.Attribute("Target")?.Value;
            if (string.IsNullOrWhiteSpace(target))
            {
                lastAttempt = Fail($"Relationship {relationshipId} không có Target.");
                continue;
            }

            var imagePath = ResolveRelationshipTarget(numberingPath, target);

            if (!package.BinaryParts.TryGetValue(imagePath, out var imageBytes))
            {
                lastAttempt = Fail($"Không đọc được ảnh {imagePath} trong file học sinh.");
                continue;
            }

            var actualHash = ComputeSha256(imageBytes);
            var isMatched = string.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase);

            return new SpecialConditionEvaluationResult
            {
                IsPassed = isMatched,
                Type = "PictureBullet",
                Message = isMatched
                    ? "Picture bullet đúng hình ảnh yêu cầu."
                    : "Picture bullet không đúng hình ảnh yêu cầu.",
                ExpectedSha256 = expectedHash,
                ActualSha256 = actualHash,
                ImagePath = imagePath,
                NumPicBulletId = lvlPicBulletId,
                RelationshipId = relationshipId
            };
        }

        return lastAttempt ?? Fail(
            config.ParagraphIndex.HasValue
                ? $"Paragraph {config.ParagraphIndex.Value} không sử dụng picture bullet phù hợp."
                : "Không tìm thấy paragraph nào sử dụng picture bullet phù hợp.");
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

private static string ComputeSha256(byte[] bytes)
{
    using var sha256 = System.Security.Cryptography.SHA256.Create();
    var hash = sha256.ComputeHash(bytes);
    return Convert.ToHexString(hash).ToLowerInvariant();
}

private static string NormalizeHash(string hash)
{
    return (hash ?? string.Empty).Trim().Replace("-", "").Replace(" ", "").ToLowerInvariant();
}
        private static ExpectedMatchResult MatchExpected(
    string actualXml,
    string expectedValue,
    string compareMode)
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
                    expectedValue);
            }

            if (string.Equals(
                mode,
                XmlGradingCompareModes.XmlContainsNormalized,
                StringComparison.OrdinalIgnoreCase))
            {
                return XmlContainsNormalized(
                    actualXml,
                    expectedValue);
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

        private static ExpectedMatchResult XmlContainsNormalized(string actualXml, string expectedValue)
        {
            var normalizedActual =
                NormalizeXmlForComparison(actualXml);

            var normalizedExpected =
                NormalizeXmlForComparison(expectedValue);

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

        private static string NormalizeXmlForComparison(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

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
        private static ExpectedMatchResult XmlEquivalentWholeFile(string actualXml, string expectedValue)
        {
            try
            {
                var normalizedActual = NormalizeXmlFragment(actualXml);
                var normalizedExpected = NormalizeXmlFragment(expectedValue);
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

        private static string NormalizeXmlFragment(string xml)
        {
            var wrapped = $"<__root>{StripXmlDeclaration(xml)}</__root>";
            var document = XDocument.Parse(wrapped, LoadOptions.PreserveWhitespace);
            var normalized = string.Concat(document.Root!.Nodes().Select(NormalizeNode));
            return normalized;
        }

        private static string NormalizeNode(XNode node)
        {
            return node switch
            {
                XElement element => NormalizeElement(element),
                XCData cdata => SecurityElement.Escape(NormalizeText(cdata.Value)) ?? string.Empty,
                XText text => SecurityElement.Escape(NormalizeText(text.Value)) ?? string.Empty,
                _ => string.Empty
            };
        }

        private static string NormalizeElement(XElement element)
        {
            var name = NormalizeName(element.Name);
            var attributes = element.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => NormalizeName(attribute.Name), StringComparer.Ordinal)
                .ThenBy(attribute => attribute.Value, StringComparer.Ordinal)
                .Select(attribute => $"{NormalizeName(attribute.Name)}=\"{SecurityElement.Escape(attribute.Value) ?? string.Empty}\"");

            var attributeText = string.Join(" ", attributes);
            var openTag = string.IsNullOrWhiteSpace(attributeText) ? $"<{name}>" : $"<{name} {attributeText}>";
            var children = string.Concat(element.Nodes().Select(NormalizeNode));

            return $"{openTag}{children}</{name}>";
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
                    matches.All(match => match.IsMatched) &&
                    matches.Select(match => match.MatchIndex ?? -1).SequenceEqual(
                        matches.Select(match => match.MatchIndex ?? -1).OrderBy(index => index)),

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

                    if (task.Conditions.Count == 0)
                    {
                        result.Errors.Add($"{taskPrefix}.conditions phải có ít nhất 1 condition.");
                    }

                    var totalConditionScore = task.Conditions.Sum(condition => condition.Score);
                    if (totalConditionScore != task.MaxScore)
                    {
                        result.Errors.Add($"{taskPrefix}.conditions tổng score ({totalConditionScore}) phải bằng task.maxScore ({task.MaxScore}).");
                    }

                    foreach (var condition in task.Conditions)
                    {
                        ValidateCondition(condition, taskPrefix, result);
                    }
                }
            }

            return result;
        }

        


        private static void ValidateSpecialCondition(
    SpecialCondition? specialCondition,
    string conditionPrefix,
    XmlRuleValidationResult result)
{
    if (specialCondition == null || specialCondition.Type == SpecialConditionType.None)
    {
        return;
    }

    switch (specialCondition.Type)
    {
        case SpecialConditionType.PictureBullet:
            var pb = specialCondition.PictureBullet;

            if (pb == null)
            {
                result.Errors.Add($"{conditionPrefix}.specialCondition.pictureBullet không được null.");
                return;
            }

            if (string.IsNullOrWhiteSpace(pb.ExpectedImageSha256))
            {
                result.Errors.Add($"{conditionPrefix}.specialCondition.pictureBullet.expectedImageSha256 không được rỗng.");
            }

            if (string.IsNullOrWhiteSpace(pb.DocumentPart))
            {
                result.Errors.Add($"{conditionPrefix}.specialCondition.pictureBullet.documentPart không được rỗng.");
            }
            else if (!IsSafeSourceFile(pb.DocumentPart))
            {
                result.Errors.Add($"{conditionPrefix}.specialCondition.pictureBullet.documentPart không hợp lệ.");
            }

            if (pb.Level.HasValue && pb.Level.Value < 0)
            {
                result.Errors.Add($"{conditionPrefix}.specialCondition.pictureBullet.level phải >= 0.");
            }

            if (pb.NumId.HasValue && pb.NumId.Value < 0)
            {
                result.Errors.Add($"{conditionPrefix}.specialCondition.pictureBullet.numId phải >= 0.");
            }

            if (pb.ParagraphIndex.HasValue && pb.ParagraphIndex.Value < 0)
            {
                result.Errors.Add($"{conditionPrefix}.specialCondition.pictureBullet.paragraphIndex phải >= 0.");
            }

            break;

        default:
            result.Errors.Add($"{conditionPrefix}.specialCondition.type không được hỗ trợ: {specialCondition.Type}.");
            break;
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
                                        ExpectedValues = new List<string>
                                        {
                                            "<c r=\"A1\" t=\"s\"><v>0</v></c>",
                                            "<c r=\"A2\" t=\"s\"><v>1</v></c>"
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
                                        ExpectedValues = new List<string>
                                        {
                                            "<c r=\"A1\" s=\"5\" t=\"s\"><v>0</v></c>"
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
                                        ExpectedValues = new List<string>
                                        {
                                            "<c r=\"A2\" s=\"6\" t=\"s\"><v>1</v></c>"
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
                                        ExpectedValues = new List<string>
                                        {
                                            "<sheetData>"
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