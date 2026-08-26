using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Text.Json.Nodes;

namespace MOS.ExcelGrading.Core.Services
{
    public class GradingConfigService : IGradingConfigService
    {
        private readonly IMongoCollection<GradingConfig> _configs;
        private readonly IMongoCollection<GradingConfigVersion> _versions;
        private readonly IMongoCollection<GradingConfigTestRun> _testRuns;
        private readonly IGradingService _gradingService;
        private readonly ILogger<GradingConfigService> _logger;

        public GradingConfigService(
            IMongoDatabase database,
            IGradingService gradingService,
            ILogger<GradingConfigService> logger)
        {
            _configs = database.GetCollection<GradingConfig>("gradingConfigs");
            _versions = database.GetCollection<GradingConfigVersion>("gradingConfigVersions");
            _testRuns = database.GetCollection<GradingConfigTestRun>("gradingConfigTestRuns");
            _gradingService = gradingService;
            _logger = logger;

            EnsureIndexes();
        }

        public async Task<List<GradingConfigListItemDto>> GetAllAsync(string? subject = null, string? status = null)
        {
            var filter = Builders<GradingConfig>.Filter.Empty;

            if (!string.IsNullOrWhiteSpace(subject))
            {
                var normalizedSubject = AssignmentFileSubjects.Normalize(subject);
                filter = Builders<GradingConfig>.Filter.And(
                    filter,
                    Builders<GradingConfig>.Filter.Eq(config => config.Subject, normalizedSubject));
            }

            if (!string.IsNullOrWhiteSpace(status))
            {
                var normalizedStatus = NormalizeStatus(status);
                filter = Builders<GradingConfig>.Filter.And(
                    filter,
                    Builders<GradingConfig>.Filter.Eq(config => config.Status, normalizedStatus));
            }

            var configs = await _configs
                .Find(filter)
                .SortBy(config => config.Subject)
                .ThenBy(config => config.ProjectNumber)
                .ThenByDescending(config => config.Version)
                .ToListAsync();

            return configs.Select(ToListItemDto).ToList();
        }

        public async Task<GradingConfigDetailDto?> GetByIdAsync(string id)
        {
            EnsureValidObjectId(id, "Mã grading config");
            var config = await _configs.Find(x => x.Id == id.Trim()).FirstOrDefaultAsync();
            return config == null ? null : ToDetailDto(config);
        }

        public async Task<GradingConfigDetailDto?> GetActiveByEndpointAsync(string gradingApiEndpoint)
        {
            var normalizedEndpoint = NormalizeEndpoint(gradingApiEndpoint);
            var config = await _configs
                .Find(x => x.GradingApiEndpoint == normalizedEndpoint && x.Status == GradingConfigStatuses.Active)
                .SortByDescending(x => x.Version)
                .FirstOrDefaultAsync();

            return config == null ? null : ToDetailDto(config);
        }

        public async Task<GradingConfigDetailDto> ImportFromCodeAsync(ImportGradingConfigFromCodeRequest request, string userId)
        {
            EnsureValidObjectId(userId, "Người dùng");
            var normalizedEndpoint = NormalizeEndpoint(request?.GradingApiEndpoint);
            EnsureEndpointParts(normalizedEndpoint, out var subject, out var projectNumber);

            var taskSnapshot = _gradingService.GetTaskSnapshotForEndpoint(normalizedEndpoint);
            if (taskSnapshot.Count == 0)
            {
                throw new InvalidOperationException("Không tìm thấy task snapshot từ hardcoded grader cho endpoint này.");
            }

            var existing = await _configs
                .Find(x => x.GradingApiEndpoint == normalizedEndpoint)
                .SortByDescending(x => x.Version)
                .FirstOrDefaultAsync();

            var now = DateTime.UtcNow;
            var config = new GradingConfig
            {
                GradingApiEndpoint = normalizedEndpoint,
                Subject = subject,
                ProjectNumber = projectNumber,
                ProjectCode = $"project{projectNumber:00}",
                DisplayName = NormalizeDisplayName(request?.DisplayName, subject, projectNumber),
                Status = request?.Publish == true ? GradingConfigStatuses.Active : GradingConfigStatuses.Draft,
                Version = (existing?.Version ?? 0) + 1,
                NormalizedMaxScore = GradingConfigDefaults.NormalizedMaxScore,
                Tasks = taskSnapshot
                    .Select((task, index) => new GradingConfigTask
                    {
                        TaskId = NormalizeRequiredText(task.TaskId, 120, "Task ID"),
                        TaskName = NormalizeRequiredText(task.TaskName, 300, "Task name"),
                        MaxScore = Convert.ToDecimal(task.MaxScore),
                        Enabled = true,
                        SortOrder = index + 1,
                        RuleType = GradingRuleTypes.LegacyCode,
                        Rule = new BsonDocument
                        {
                            ["gradingApiEndpoint"] = normalizedEndpoint,
                            ["taskId"] = NormalizeRequiredText(task.TaskId, 120, "Task ID")
                        }
                    })
                    .ToList(),
                CreatedAt = now,
                CreatedBy = userId.Trim(),
                UpdatedAt = now,
                UpdatedBy = userId.Trim(),
                PublishedAt = request?.Publish == true ? now : null,
                PublishedBy = request?.Publish == true ? userId.Trim() : null,
                Summary = NormalizeOptionalText(request?.Summary, 1000)
            };

            if (request?.Publish == true)
            {
                await DeactivateExistingActiveAsync(normalizedEndpoint);
            }

            await _configs.InsertOneAsync(config);
            await CreateVersionAsync(config, GradingConfigVersionActions.ImportFromCode, config.Summary, userId);

            _logger.LogInformation("✅ Imported grading config {ConfigId} for {Endpoint}", config.Id, normalizedEndpoint);
            return ToDetailDto(config);
        }

        public async Task<GradingConfigDetailDto> UpdateAsync(string id, UpdateGradingConfigRequest request, string userId)
        {
            EnsureValidObjectId(id, "Mã grading config");
            EnsureValidObjectId(userId, "Người dùng");

            var config = await _configs.Find(x => x.Id == id.Trim()).FirstOrDefaultAsync();
            if (config == null)
            {
                throw new KeyNotFoundException("Không tìm thấy grading config.");
            }

            if (config.Status == GradingConfigStatuses.Active)
            {
                throw new InvalidOperationException("Không chỉnh trực tiếp config active. Hãy restore/import thành bản draft mới rồi publish.");
            }

            config.DisplayName = NormalizeDisplayName(request?.DisplayName ?? config.DisplayName, config.Subject, config.ProjectNumber);
            config.Summary = NormalizeOptionalText(request?.Summary, 1000);
            config.Tasks = NormalizeTasks(request?.Tasks);
            config.UpdatedAt = DateTime.UtcNow;
            config.UpdatedBy = userId.Trim();

            await _configs.ReplaceOneAsync(x => x.Id == config.Id, config);
            return ToDetailDto(config);
        }

        public async Task<GradingConfigDetailDto> PublishAsync(string id, PublishGradingConfigRequest request, string userId)
        {
            EnsureValidObjectId(id, "Mã grading config");
            EnsureValidObjectId(userId, "Người dùng");

            var config = await _configs.Find(x => x.Id == id.Trim()).FirstOrDefaultAsync();
            if (config == null)
            {
                throw new KeyNotFoundException("Không tìm thấy grading config.");
            }

            if (!config.Tasks.Any(x => x.Enabled))
            {
                throw new InvalidOperationException("Config phải có ít nhất một task đang bật.");
            }

            await DeactivateExistingActiveAsync(config.GradingApiEndpoint, config.Id);

            config.Status = GradingConfigStatuses.Active;
            config.Version += 1;
            config.Summary = NormalizeOptionalText(request?.Summary, 1000) ?? config.Summary;
            config.UpdatedAt = DateTime.UtcNow;
            config.UpdatedBy = userId.Trim();
            config.PublishedAt = DateTime.UtcNow;
            config.PublishedBy = userId.Trim();

            await _configs.ReplaceOneAsync(x => x.Id == config.Id, config);
            await CreateVersionAsync(config, GradingConfigVersionActions.Publish, config.Summary, userId);

            return ToDetailDto(config);
        }

        public async Task<List<GradingConfigVersionDto>> GetVersionsAsync(string id)
        {
            EnsureValidObjectId(id, "Mã grading config");

            var versions = await _versions
                .Find(x => x.GradingConfigId == id.Trim())
                .SortByDescending(x => x.Version)
                .ToListAsync();

            return versions.Select(ToVersionDto).ToList();
        }

        public async Task<GradingConfigDetailDto?> GetVersionSnapshotAsync(string id, int version)
        {
            EnsureValidObjectId(id, "Mã grading config");

            var item = await _versions
                .Find(x => x.GradingConfigId == id.Trim() && x.Version == version)
                .FirstOrDefaultAsync();

            return item == null ? null : ToDetailDto(item.Snapshot);
        }

        public async Task<GradingConfigDetailDto> RestoreVersionAsync(string id, int version, RestoreGradingConfigVersionRequest request, string userId)
        {
            EnsureValidObjectId(id, "Mã grading config");
            EnsureValidObjectId(userId, "Người dùng");

            var current = await _configs.Find(x => x.Id == id.Trim()).FirstOrDefaultAsync();
            if (current == null)
            {
                throw new KeyNotFoundException("Không tìm thấy grading config.");
            }

            var versionItem = await _versions
                .Find(x => x.GradingConfigId == id.Trim() && x.Version == version)
                .FirstOrDefaultAsync();

            if (versionItem == null)
            {
                throw new KeyNotFoundException("Không tìm thấy version grading config.");
            }

            var restored = versionItem.Snapshot;
            restored.Id = ObjectId.GenerateNewId().ToString();
            restored.Status = GradingConfigStatuses.Draft;
            restored.Version = current.Version + 1;
            restored.CreatedAt = DateTime.UtcNow;
            restored.CreatedBy = userId.Trim();
            restored.UpdatedAt = DateTime.UtcNow;
            restored.UpdatedBy = userId.Trim();
            restored.PublishedAt = null;
            restored.PublishedBy = null;
            restored.Summary = NormalizeOptionalText(request?.Summary, 1000) ?? $"Restore từ version {version}";

            await _configs.InsertOneAsync(restored);
            await CreateVersionAsync(restored, GradingConfigVersionActions.Restore, restored.Summary, userId);

            return ToDetailDto(restored);
        }

        public async Task<List<GradingConfigTestRunDto>> GetTestRunsAsync(string id)
        {
            EnsureValidObjectId(id, "Mã grading config");

            var testRuns = await _testRuns
                .Find(x => x.GradingConfigId == id.Trim())
                .SortByDescending(x => x.CreatedAt)
                .Limit(50)
                .ToListAsync();

            return testRuns.Select(ToTestRunDto).ToList();
        }

        public async Task<GradingConfigTestRunDto> CreateTestRunAsync(
            string id,
            GradingResult result,
            string fileName,
            bool usedOverride,
            string userId,
            string? error = null)
        {
            EnsureValidObjectId(id, "Mã grading config");
            EnsureValidObjectId(userId, "Người dùng");

            var config = await _configs.Find(x => x.Id == id.Trim()).FirstOrDefaultAsync();
            if (config == null)
            {
                throw new KeyNotFoundException("Không tìm thấy grading config.");
            }

            var testRun = new GradingConfigTestRun
            {
                GradingConfigId = config.Id,
                GradingApiEndpoint = config.GradingApiEndpoint,
                ConfigVersion = config.Version,
                UsedOverride = usedOverride,
                FileName = NormalizeOptionalText(fileName, 260) ?? string.Empty,
                TotalScore = result.TotalScore,
                MaxScore = result.MaxScore,
                Percentage = result.Percentage,
                Status = result.Status,
                CreatedAt = DateTime.UtcNow,
                CreatedBy = userId.Trim(),
                Error = NormalizeOptionalText(error, 1000)
            };

            await _testRuns.InsertOneAsync(testRun);
            return ToTestRunDto(testRun);
        }

        public Task<List<GradingRuleTypeDto>> GetRuleTypesAsync()
        {
            return Task.FromResult(new List<GradingRuleTypeDto>
            {
                new()
                {
                    RuleType = GradingRuleTypes.LegacyCode,
                    DisplayName = "Legacy hardcoded grader",
                    Description = "V1 giữ engine chấm hardcoded hiện tại; rule chỉ ghi taskId/endpoint để phục vụ quản trị version."
                }
            });
        }

        private void EnsureIndexes()
        {
            try
            {
                _configs.Indexes.CreateMany(new[]
                {
                    new CreateIndexModel<GradingConfig>(
                        Builders<GradingConfig>.IndexKeys
                            .Ascending(x => x.GradingApiEndpoint)
                            .Ascending(x => x.Status)),
                    new CreateIndexModel<GradingConfig>(
                        Builders<GradingConfig>.IndexKeys
                            .Ascending(x => x.Subject)
                            .Ascending(x => x.ProjectNumber))
                });

                _versions.Indexes.CreateMany(new[]
                {
                    new CreateIndexModel<GradingConfigVersion>(
                        Builders<GradingConfigVersion>.IndexKeys
                            .Ascending(x => x.GradingConfigId)
                            .Descending(x => x.Version))
                });

                _testRuns.Indexes.CreateMany(new[]
                {
                    new CreateIndexModel<GradingConfigTestRun>(
                        Builders<GradingConfigTestRun>.IndexKeys
                            .Ascending(x => x.GradingConfigId)
                            .Descending(x => x.CreatedAt))
                });
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "⚠️ Unable to ensure indexes for grading config collections");
            }
        }

        private async Task DeactivateExistingActiveAsync(string endpoint, string? exceptId = null)
        {
            var filter = Builders<GradingConfig>.Filter.And(
                Builders<GradingConfig>.Filter.Eq(x => x.GradingApiEndpoint, endpoint),
                Builders<GradingConfig>.Filter.Eq(x => x.Status, GradingConfigStatuses.Active));

            if (!string.IsNullOrWhiteSpace(exceptId))
            {
                filter = Builders<GradingConfig>.Filter.And(
                    filter,
                    Builders<GradingConfig>.Filter.Ne(x => x.Id, exceptId));
            }

            var update = Builders<GradingConfig>.Update
                .Set(x => x.Status, GradingConfigStatuses.Draft)
                .Set(x => x.UpdatedAt, DateTime.UtcNow);

            await _configs.UpdateManyAsync(filter, update);
        }

        private async Task CreateVersionAsync(GradingConfig config, string action, string? summary, string userId)
        {
            var version = new GradingConfigVersion
            {
                GradingConfigId = config.Id,
                Version = config.Version,
                Action = action,
                Summary = summary,
                Snapshot = CloneForSnapshot(config),
                CreatedAt = DateTime.UtcNow,
                CreatedBy = userId.Trim()
            };

            await _versions.InsertOneAsync(version);
        }

        private static GradingConfig CloneForSnapshot(GradingConfig config)
        {
            return new GradingConfig
            {
                Id = config.Id,
                GradingApiEndpoint = config.GradingApiEndpoint,
                Subject = config.Subject,
                ProjectNumber = config.ProjectNumber,
                ProjectCode = config.ProjectCode,
                DisplayName = config.DisplayName,
                Status = config.Status,
                Version = config.Version,
                NormalizedMaxScore = config.NormalizedMaxScore,
                Tasks = config.Tasks.Select(task => new GradingConfigTask
                {
                    TaskId = task.TaskId,
                    TaskName = task.TaskName,
                    MaxScore = task.MaxScore,
                    Enabled = task.Enabled,
                    SortOrder = task.SortOrder,
                    RuleType = task.RuleType,
                    Rule = new BsonDocument(task.Rule)
                }).ToList(),
                CreatedAt = config.CreatedAt,
                CreatedBy = config.CreatedBy,
                UpdatedAt = config.UpdatedAt,
                UpdatedBy = config.UpdatedBy,
                PublishedAt = config.PublishedAt,
                PublishedBy = config.PublishedBy,
                Summary = config.Summary
            };
        }

        private static List<GradingConfigTask> NormalizeTasks(List<UpdateGradingConfigTaskRequest>? tasks)
        {
            if (tasks == null || tasks.Count == 0)
            {
                throw new InvalidOperationException("Config phải có ít nhất một task.");
            }

            return tasks
                .Select((task, index) => new GradingConfigTask
                {
                    TaskId = NormalizeRequiredText(task.TaskId, 120, "Task ID"),
                    TaskName = NormalizeRequiredText(task.TaskName, 300, "Task name"),
                    MaxScore = task.MaxScore > 0 ? task.MaxScore : throw new InvalidOperationException("Max score của task phải lớn hơn 0."),
                    Enabled = task.Enabled,
                    SortOrder = task.SortOrder > 0 ? task.SortOrder : index + 1,
                    RuleType = NormalizeRuleType(task.RuleType),
                    Rule = JsonNodeToBsonDocument(task.Rule)
                })
                .OrderBy(task => task.SortOrder)
                .ToList();
        }

        private static string NormalizeEndpoint(string? endpoint)
        {
            var normalized = GradingApiEndpoints.NormalizeEndpoint(endpoint);
            if (string.IsNullOrWhiteSpace(normalized) || !GradingApiEndpoints.IsValidEndpoint(normalized))
            {
                throw new InvalidOperationException("Grading API endpoint không hợp lệ.");
            }

            return normalized;
        }

        private static void EnsureEndpointParts(string endpoint, out string subject, out int projectNumber)
        {
            if (!GradingApiEndpoints.TryExtractSubject(endpoint, out subject) ||
                !GradingApiEndpoints.TryExtractProjectNumber(endpoint, out projectNumber))
            {
                throw new InvalidOperationException("Không xác định được subject/project number từ endpoint.");
            }
        }

        private static string NormalizeStatus(string status)
        {
            var normalized = status.Trim().ToLowerInvariant();
            if (!GradingConfigStatuses.IsValid(normalized))
            {
                throw new InvalidOperationException("Trạng thái grading config không hợp lệ.");
            }

            return normalized;
        }

        private static string NormalizeRuleType(string? ruleType)
        {
            var normalized = string.IsNullOrWhiteSpace(ruleType)
                ? GradingRuleTypes.LegacyCode
                : ruleType.Trim().ToLowerInvariant();

            if (normalized != GradingRuleTypes.LegacyCode)
            {
                throw new InvalidOperationException("V1 chỉ hỗ trợ ruleType legacy.code.");
            }

            return normalized;
        }

        private static string NormalizeDisplayName(string? value, string subject, int projectNumber)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                var trimmed = value.Trim();
                return trimmed.Length <= 200 ? trimmed : trimmed[..200];
            }

            return $"Project {projectNumber:00} - {subject.ToUpperInvariant()}";
        }

        private static string NormalizeRequiredText(string? value, int maxLength, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException($"{fieldName} không được để trống.");
            }

            var trimmed = value.Trim();
            return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
        }

        private static string? NormalizeOptionalText(string? value, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var trimmed = value.Trim();
            return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
        }

        private static void EnsureValidObjectId(string? value, string label)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new InvalidOperationException($"{label} không hợp lệ.");
            }
        }

        private static BsonDocument JsonNodeToBsonDocument(JsonNode? node)
        {
            if (node == null)
            {
                return new BsonDocument();
            }

            return BsonDocument.Parse(node.ToJsonString());
        }

        private static JsonNode? BsonDocumentToJsonNode(BsonDocument document)
        {
            if (document == null || document.ElementCount == 0)
            {
                return null;
            }

            return JsonNode.Parse(document.ToJson());
        }

        private static GradingConfigListItemDto ToListItemDto(GradingConfig config)
        {
            var enabledTasks = config.Tasks.Where(task => task.Enabled).ToList();

            return new GradingConfigListItemDto
            {
                Id = config.Id,
                GradingApiEndpoint = config.GradingApiEndpoint,
                Subject = config.Subject,
                ProjectNumber = config.ProjectNumber,
                ProjectCode = config.ProjectCode,
                DisplayName = config.DisplayName,
                Status = config.Status,
                Version = config.Version,
                NormalizedMaxScore = config.NormalizedMaxScore,
                TaskCount = config.Tasks.Count,
                EnabledTaskCount = enabledTasks.Count,
                RawMaxScore = enabledTasks.Sum(task => task.MaxScore),
                CreatedAt = config.CreatedAt,
                UpdatedAt = config.UpdatedAt,
                PublishedAt = config.PublishedAt,
                Summary = config.Summary
            };
        }

        private static GradingConfigDetailDto ToDetailDto(GradingConfig config)
        {
            var item = ToListItemDto(config);

            return new GradingConfigDetailDto
            {
                Id = item.Id,
                GradingApiEndpoint = item.GradingApiEndpoint,
                Subject = item.Subject,
                ProjectNumber = item.ProjectNumber,
                ProjectCode = item.ProjectCode,
                DisplayName = item.DisplayName,
                Status = item.Status,
                Version = item.Version,
                NormalizedMaxScore = item.NormalizedMaxScore,
                TaskCount = item.TaskCount,
                EnabledTaskCount = item.EnabledTaskCount,
                RawMaxScore = item.RawMaxScore,
                CreatedAt = item.CreatedAt,
                UpdatedAt = item.UpdatedAt,
                PublishedAt = item.PublishedAt,
                Summary = item.Summary,
                Tasks = config.Tasks
                    .OrderBy(task => task.SortOrder)
                    .Select(task => new GradingConfigTaskDto
                    {
                        TaskId = task.TaskId,
                        TaskName = task.TaskName,
                        MaxScore = task.MaxScore,
                        Enabled = task.Enabled,
                        SortOrder = task.SortOrder,
                        RuleType = task.RuleType,
                        Rule = BsonDocumentToJsonNode(task.Rule)
                    })
                    .ToList()
            };
        }

        private static GradingConfigTestRunDto ToTestRunDto(GradingConfigTestRun testRun)
        {
            return new GradingConfigTestRunDto
            {
                Id = testRun.Id,
                GradingConfigId = testRun.GradingConfigId,
                GradingApiEndpoint = testRun.GradingApiEndpoint,
                ConfigVersion = testRun.ConfigVersion,
                UsedOverride = testRun.UsedOverride,
                FileName = testRun.FileName,
                TotalScore = testRun.TotalScore,
                MaxScore = testRun.MaxScore,
                Percentage = testRun.Percentage,
                Status = testRun.Status,
                CreatedAt = testRun.CreatedAt,
                CreatedBy = testRun.CreatedBy,
                Error = testRun.Error
            };
        }

        private static GradingConfigVersionDto ToVersionDto(GradingConfigVersion version)
        {
            return new GradingConfigVersionDto
            {
                Id = version.Id,
                GradingConfigId = version.GradingConfigId,
                Version = version.Version,
                Action = version.Action,
                Summary = version.Summary,
                CreatedAt = version.CreatedAt,
                CreatedBy = version.CreatedBy
            };
        }
    }
}