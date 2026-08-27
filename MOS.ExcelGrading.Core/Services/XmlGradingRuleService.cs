using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;
using System.Xml.Linq;

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

            var packageParts = ReadPackageXmlParts(studentFile);

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
                    var conditionResult = EvaluateCondition(condition, packageParts);
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

        private static Dictionary<string, string> ReadPackageXmlParts(Stream studentFile)
        {
            if (studentFile.CanSeek)
            {
                studentFile.Position = 0;
            }

            using var archive = new ZipArchive(studentFile, ZipArchiveMode.Read, leaveOpen: true);
            var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    using var entryStream = entry.Open();
                    using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    parts[NormalizeSourceFile(entry.FullName)] = reader.ReadToEnd();
                }
            }

            return parts;
        }

        private static XmlConditionEvaluationResult EvaluateCondition(
            XmlGradingCondition condition,
            IReadOnlyDictionary<string, string> packageParts)
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

            if (!packageParts.TryGetValue(result.SourceFile, out var actualXml))
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

        private static ExpectedMatchResult MatchExpected(string actualXml, string expectedValue, string compareMode)
        {
            var mode = string.IsNullOrWhiteSpace(compareMode)
                ? XmlGradingCompareModes.XmlContainsNormalized
                : compareMode.Trim();

            return mode switch
            {
                var value when string.Equals(value, XmlGradingCompareModes.ExactStringContains, StringComparison.OrdinalIgnoreCase) =>
                    RawContains(actualXml, expectedValue, trim: false),

                var value when string.Equals(value, XmlGradingCompareModes.XmlContains, StringComparison.OrdinalIgnoreCase) =>
                    RawContains(actualXml, expectedValue, trim: true),

                var value when string.Equals(value, XmlGradingCompareModes.XmlEquivalentWholeFile, StringComparison.OrdinalIgnoreCase) =>
                    XmlEquivalentWholeFile(actualXml, expectedValue),

                _ => XmlContainsNormalized(actualXml, expectedValue)
            };
        }

        private static ExpectedMatchResult RawContains(string actualXml, string expectedValue, bool trim)
        {
            var expected = trim ? expectedValue.Trim() : expectedValue;
            var index = actualXml.IndexOf(expected, StringComparison.Ordinal);
            return new ExpectedMatchResult
            {
                ExpectedValue = expectedValue,
                IsMatched = index >= 0,
                MatchIndex = index >= 0 ? index : null
            };
        }

        private static ExpectedMatchResult XmlContainsNormalized(string actualXml, string expectedValue)
        {
            try
            {
                var normalizedActual = NormalizeXmlFragment(actualXml);
                var normalizedExpected = NormalizeXmlFragment(expectedValue);
                var index = normalizedActual.IndexOf(normalizedExpected, StringComparison.Ordinal);
                return new ExpectedMatchResult
                {
                    ExpectedValue = expectedValue,
                    IsMatched = index >= 0,
                    MatchIndex = index >= 0 ? index : null
                };
            }
            catch (XmlException)
            {
                return RawContains(actualXml, expectedValue, trim: true);
            }
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

            if (condition.ExpectedValues.Count == 0 || condition.ExpectedValues.Any(string.IsNullOrWhiteSpace))
            {
                result.Errors.Add($"{conditionPrefix}.expectedValue phải là string không rỗng hoặc array string không rỗng.");
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
        }

        private static bool IsSafeSourceFile(string sourceFile)
        {
            var normalized = NormalizeSourceFile(sourceFile);
            return !string.IsNullOrWhiteSpace(normalized) &&
                   !normalized.Contains("..", StringComparison.Ordinal) &&
                   !Path.IsPathRooted(normalized) &&
                   normalized.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeSourceFile(string sourceFile)
        {
            return (sourceFile ?? string.Empty).Trim().Replace('\\', '/').TrimStart('/');
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