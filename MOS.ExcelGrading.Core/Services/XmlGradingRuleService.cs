using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
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
                    var specialResult = EvaluateTaskSpecialCondition(taskRule.SpecialCondition!, package);
                    specialConditionPassed = specialResult.IsPassed;

                    if (specialResult.IsPassed)
                    {
                        taskResult.Details.Add($"[SpecialCondition:{taskRule.SpecialCondition!.Type}] {specialResult.Message}");

                        // Cộng điểm riêng của specialCondition do người tạo ruleset cấu hình.
                        // Cho phép Task chỉ dùng specialCondition (0 condition XML) hoặc
                        // kết hợp cả hai, miễn tổng = task.maxScore (được validate ở ValidateRuleSet).
                        taskResult.Score += taskRule.SpecialCondition!.Score;
                    }
                    else
                    {
                        taskResult.Errors.Add($"[SpecialCondition:{taskRule.SpecialCondition!.Type}] {specialResult.Message}");
                    }
                }

                // ===== NORMAL XML CONDITIONS =====
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

                // Special condition FAIL -> zero toàn bộ Task, kể cả khi có conditions XML đã đạt điểm.
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
                }
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

            if (condition.ExpectedValues.Count == 0)
            {
                throw new InvalidOperationException("expectedValue là bắt buộc.");
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
            result.MatchedExpectedValues = matches.Where(match => match.IsMatched).Select(match => match.ExpectedValue).ToList();
            result.MissingExpectedValues = matches.Where(match => !match.IsMatched).Select(match => match.ExpectedValue).ToList();

            return result;
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

        private static SpecialConditionEvalOutcome EvaluateTaskSpecialCondition(SpecialCondition specialCondition, OfficePackage package)
        {
            if (string.Equals(specialCondition.Type, SpecialConditionTypes.PictureBullet, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluatePictureBullet(specialCondition.Config, package);
            }

            // MỚI
            if (string.Equals(specialCondition.Type, SpecialConditionTypes.InsertedImage, StringComparison.OrdinalIgnoreCase))
            {
                return EvaluateInsertedImage(specialCondition.ImageInsertConfig, package);
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

            if (!package.XmlParts.TryGetValue(documentPart, out var documentXml))
            {
                return Fail($"Không tìm thấy {documentPart} trong file học sinh.");
            }

            if (!package.XmlParts.TryGetValue(documentRelsPath, out var relsXml))
            {
                return Fail("Không tìm thấy word/_rels/document.xml.rels.");
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

            try
            {
                var documentDocument = XDocument.Parse(documentXml);
                var relsDocument = XDocument.Parse(relsXml);
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

                    var relationship = relsDocument
                        .Descendants(rel + "Relationship")
                        .FirstOrDefault(e => string.Equals(e.Attribute("Id")?.Value, relationshipId, StringComparison.Ordinal));

                    var target = relationship?.Attribute("Target")?.Value;
                    if (string.IsNullOrWhiteSpace(target))
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

                    var actualHash = ImageHashUtility.ComputeSha256(imageBytes);

                    if (!IsImageMatch(imageBytes, actualHash, expectedHash, expectedPerceptualHash))
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

            if (!package.XmlParts.TryGetValue(documentPart, out var documentXml))
            {
                return Fail($"Không tìm thấy {documentPart} trong file học sinh.");
            }

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

                var expectedHash = ImageHashUtility.NormalizeHash(config.ImageHash);
                var expectedPerceptualHash = config.PerceptualHash;

                // Duyệt toàn bộ paragraph trong document.xml, tìm bất kỳ paragraph nào
                // dùng picture bullet ở đúng Level (nếu có chỉ định) mà ảnh khớp expectedHash.
                var allParagraphs = documentDocument.Descendants(w + "p").ToList();

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
                    var numElement = numberingDocument
                        .Descendants(w + "num")
                        .FirstOrDefault(e => e.Attribute(w + "numId")?.Value == numId.ToString());

                    if (numElement == null)
                    {
                        continue;
                    }

                    var abstractNumIdVal = numElement.Element(w + "abstractNumId")?.Attribute(w + "val")?.Value;
                    if (!int.TryParse(abstractNumIdVal, out var abstractNumId))
                    {
                        continue;
                    }

                    // abstractNumId -> abstractNum -> lvl[ilvl] -> lvlPicBulletId
                    var abstractNumElement = numberingDocument
                        .Descendants(w + "abstractNum")
                        .FirstOrDefault(e => e.Attribute(w + "abstractNumId")?.Value == abstractNumId.ToString());

                    if (abstractNumElement == null)
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
                    var numPicBulletElement = numberingDocument
                        .Descendants(w + "numPicBullet")
                        .FirstOrDefault(e => e.Attribute(w + "numPicBulletId")?.Value == lvlPicBulletId.ToString());

                    if (numPicBulletElement == null)
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

                    var relationship = relsDocument
                        .Descendants(rel + "Relationship")
                        .FirstOrDefault(e => string.Equals(
                            e.Attribute("Id")?.Value, relationshipId, StringComparison.Ordinal));

                    var target = relationship?.Attribute("Target")?.Value;
                    if (string.IsNullOrWhiteSpace(target))
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

                    var actualHash = ImageHashUtility.ComputeSha256(imageBytes);

                    if (IsImageMatch(imageBytes, actualHash, expectedHash, expectedPerceptualHash))
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
                        ValidateTaskSpecialCondition(task.SpecialCondition!, taskPrefix, result);
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

            // if (condition.ExpectedValues.Count == 0 || condition.ExpectedValues.Any(string.IsNullOrWhiteSpace))
            // {
            //     result.Errors.Add($"{conditionPrefix}.expectedValue phải là string không rỗng hoặc array string không rỗng.");
            // }

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
        }

        private static void ValidateTaskSpecialCondition(
            SpecialCondition specialCondition,
            string taskPrefix,
            XmlRuleValidationResult result)
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