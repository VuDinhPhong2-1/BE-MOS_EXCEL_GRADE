// MOS.ExcelGrading.Core/Services/ScoreService.cs
using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Text.RegularExpressions;

namespace MOS.ExcelGrading.Core.Services
{
    public class ScoreService : IScoreService
    {
        private readonly IMongoCollection<Score> _scores;
        private readonly IMongoCollection<BsonDocument> _scoreDocuments;
        private readonly IMongoCollection<BsonDocument> _assignmentDocuments;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<Assignment> _assignments;
        private readonly IMongoCollection<User> _users;
        private readonly IAnalyticsService _analyticsService;
        private readonly ILogger<ScoreService> _logger;
        private static readonly Regex AutoErrorQuestionRegex = new(
            @"^\s*cau\s*(?<question>\d+)\s*(?:-\s*(?<taskName>[^:]+))?\s*:\s*(?<message>.+)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public ScoreService(
            IMongoDatabase database,
            IAnalyticsService analyticsService,
            ILogger<ScoreService> logger)
        {
            _scores = database.GetCollection<Score>("scores");
            _scoreDocuments = database.GetCollection<BsonDocument>("scores");
            _assignmentDocuments = database.GetCollection<BsonDocument>("assignments");
            _students = database.GetCollection<Student>("students");
            _assignments = database.GetCollection<Assignment>("assignments");
            _users = database.GetCollection<User>("users");
            _analyticsService = analyticsService;
            _logger = logger;
        }

        public async Task<List<ScoreResponse>> GetScoresByAssignmentAsync(string assignmentId)
        {
            try
            {
                EnsureValidObjectId(assignmentId, "Mã bài tập");

                var scores = await _scores.Find(s => s.AssignmentId == assignmentId).ToListAsync();
                var assignment = await SafeGetAssignmentAsync(assignmentId);
                return await BuildScoreResponsesAsync(scores, assignment);
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(
                    ex,
                    "⚠️ Deserialization error while reading scores for assignment {AssignmentId}. Falling back to tolerant document mapping.",
                    assignmentId);

                return await GetScoresByAssignmentFromRawDocumentsAsync(assignmentId);
            }
            catch (FormatException ex)
            {
                _logger.LogWarning(
                    ex,
                    "⚠️ Invalid score document format while reading scores for assignment {AssignmentId}. Falling back to tolerant document mapping.",
                    assignmentId);

                return await GetScoresByAssignmentFromRawDocumentsAsync(assignmentId);
            }
            catch (Exception ex)
            {
                var diagnostic = await BuildScoreSchemaDiagnosticAsync(assignmentId);
                _logger.LogError(
                    ex,
                    "❌ Error getting scores for assignment {AssignmentId}. Diagnostic: {Diagnostic}",
                    assignmentId,
                    diagnostic);
                throw new InvalidOperationException(
                    $"Không thể đọc điểm của bài tập. {diagnostic}",
                    ex);
            }
        }

        public async Task<List<ScoreResponse>> GetScoresByStudentAsync(string studentId)
        {
            try
            {
                var scores = await _scores.Find(s => s.StudentId == studentId).ToListAsync();
                var result = new List<ScoreResponse>();

                var student = await _students.Find(s => s.Id == studentId).FirstOrDefaultAsync();

                foreach (var score in scores)
                {
                    var assignment = await _assignments.Find(a => a.Id == score.AssignmentId).FirstOrDefaultAsync();

                    var graderName = await ResolveGraderNameAsync(score.GradedBy);

                    result.Add(new ScoreResponse
                    {
                        Id = score.Id,
                        StudentId = score.StudentId,
                        StudentFirstName = student?.FirstName ?? "",
                        StudentMiddleName = student?.MiddleName ?? "",
                        StudentFullName = $"{student?.MiddleName} {student?.FirstName}".Trim(),
                        AssignmentId = score.AssignmentId,
                        AssignmentName = assignment?.Name ?? "",
                        ScoreValue = score.ScoreValue,
                        Feedback = score.Feedback,
                        AutoGradingErrors = score.AutoGradingErrors ?? new List<string>(),
                        GradedAt = score.GradedAt,
                        GradedBy = score.GradedBy,
                        GradedByName = graderName
                    });
                }

                return result.OrderByDescending(r => r.GradedAt).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting scores for student {StudentId}", studentId);
                throw;
            }
        }

        public async Task<StudentScoreReportResponse> GetStudentScoreReportAsync(string studentId, string classId)
        {
            try
            {
                var student = await _students.Find(s => s.Id == studentId).FirstOrDefaultAsync();
                if (student == null)
                    throw new Exception($"Không tìm thấy học sinh {studentId}");

                var assignments = await _assignments.Find(a =>
                    a.ClassId == classId && a.IsActive).ToListAsync();

                var scores = await _scores.Find(s =>
                    s.StudentId == studentId && s.ClassId == classId).ToListAsync();

                var scoreDetails = new List<ScoreDetailResponse>();
                double totalScore = 0;
                int completedCount = 0;

                foreach (var assignment in assignments)
                {
                    var score = scores.FirstOrDefault(s => s.AssignmentId == assignment.Id);

                    scoreDetails.Add(new ScoreDetailResponse
                    {
                        AssignmentName = assignment.Name,
                        ScoreValue = score?.ScoreValue,
                        MaxScore = assignment.MaxScore,
                        Feedback = score?.Feedback,
                        AutoGradingErrors = score?.AutoGradingErrors ?? new List<string>(),
                        GradedAt = score?.GradedAt
                    });

                    if (score?.ScoreValue.HasValue == true)
                    {
                        totalScore += score.ScoreValue.Value;
                        completedCount++;
                    }
                }

                return new StudentScoreReportResponse
                {
                    StudentId = studentId,
                    StudentFullName = $"{student.MiddleName} {student.FirstName}".Trim(),
                    Scores = scoreDetails,
                    AverageScore = completedCount > 0 ? Math.Round(totalScore / completedCount, 2) : 0,
                    TotalAssignments = assignments.Count,
                    CompletedAssignments = completedCount
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting score report for student {StudentId}", studentId);
                throw;
            }
        }

        public async Task<Score?> GetScoreAsync(string studentId, string assignmentId)
        {
            try
            {
                return await _scores.Find(s =>
                    s.StudentId == studentId && s.AssignmentId == assignmentId
                ).FirstOrDefaultAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting score");
                throw;
            }
        }

        public async Task<Score> CreateOrUpdateScoreAsync(CreateScoreRequest request, string gradedBy)
        {
            try
            {
                ValidateCreateOrUpdateRequest(request);
                EnsureValidObjectId(gradedBy, "Người chấm điểm");

                var normalizedScoreValue = NormalizeScoreValue(request.ScoreValue);
                var normalizedFeedback = NormalizeFeedback(request.Feedback);
                var normalizedAutoErrors = NormalizeAutoGradingErrors(request.AutoGradingErrors);
                var normalizedAutoTaskResults = NormalizeAutoGradingTaskResults(request.AutoGradingTaskResults);

                await EnsureStudentCanBeGradedAsync(request.StudentId);
                var assignment = await _assignments.Find(a => a.Id == request.AssignmentId).FirstOrDefaultAsync();
                if (assignment == null)
                {
                    throw new InvalidOperationException($"Không tìm thấy bài tập: {request.AssignmentId}");
                }

                if (!string.Equals(assignment.ClassId, request.ClassId, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Bài tập không thuộc lớp đang lưu điểm.");
                }

                if (normalizedScoreValue.HasValue && normalizedScoreValue.Value > assignment.MaxScore)
                {
                    throw new InvalidOperationException(
                        $"Điểm không hợp lệ. Điểm tối đa của bài '{assignment.Name}' là {assignment.MaxScore:0.##}.");
                }

                var normalizedRequest = new CreateScoreRequest
                {
                    StudentId = request.StudentId,
                    AssignmentId = request.AssignmentId,
                    ClassId = request.ClassId,
                    ScoreValue = normalizedScoreValue,
                    Feedback = normalizedFeedback,
                    AutoGradingErrors = normalizedAutoErrors,
                    AutoGradingTaskResults = normalizedAutoTaskResults
                };

                var existingScore = await GetScoreAsync(request.StudentId, request.AssignmentId);

                if (existingScore != null)
                {
                    // Update existing score
                    var update = Builders<Score>.Update
                        .Set(s => s.ScoreValue, normalizedScoreValue)
                        .Set(s => s.Feedback, normalizedFeedback)
                        .Set(s => s.AutoGradingErrors, normalizedAutoErrors)
                        .Set(s => s.GradedAt, DateTime.UtcNow)
                        .Set(s => s.GradedBy, gradedBy)
                        .Set(s => s.UpdatedAt, DateTime.UtcNow)
                        .Set(s => s.UpdatedBy, gradedBy)
                        // Làm sạch dữ liệu metadata cũ để tránh lỗi deserialize ở các bản ghi legacy.
                        .Set(s => s.CreatedBy, gradedBy);

                    await _scores.UpdateOneAsync(s => s.Id == existingScore.Id, update);

                    existingScore.ScoreValue = normalizedScoreValue;
                    existingScore.Feedback = normalizedFeedback;
                    existingScore.AutoGradingErrors = normalizedAutoErrors;
                    existingScore.GradedAt = DateTime.UtcNow;
                    existingScore.GradedBy = gradedBy;
                    existingScore.UpdatedAt = DateTime.UtcNow;
                    existingScore.UpdatedBy = gradedBy;
                    existingScore.CreatedBy = gradedBy;

                    _logger.LogInformation("✅ Score updated for student {StudentId} by {GradedBy}",
                        request.StudentId, gradedBy);

                    await SaveGradingAttemptFromSavedScoreAsync(normalizedRequest, assignment, gradedBy);

                    return existingScore;
                }
                else
                {
                    // Create new score
                    var score = new Score
                    {
                        StudentId = request.StudentId,
                        AssignmentId = request.AssignmentId,
                        ClassId = request.ClassId,
                        ScoreValue = normalizedScoreValue,
                        Feedback = normalizedFeedback,
                        AutoGradingErrors = normalizedAutoErrors,
                        GradedAt = DateTime.UtcNow,
                        GradedBy = gradedBy,
                        CreatedBy = gradedBy,
                        CreatedAt = DateTime.UtcNow
                    };

                    await _scores.InsertOneAsync(score);

                    _logger.LogInformation("✅ Score created for student {StudentId} by {GradedBy}",
                        request.StudentId, gradedBy);

                    await SaveGradingAttemptFromSavedScoreAsync(normalizedRequest, assignment, gradedBy);

                    return score;
                }
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (FormatException ex)
            {
                _logger.LogWarning(ex, "⚠️ Invalid score payload format");
                throw new InvalidOperationException("Dữ liệu lưu điểm không hợp lệ.");
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(ex, "⚠️ Invalid bson payload when saving score");
                throw new InvalidOperationException("Dữ liệu lưu điểm không hợp lệ.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error creating/updating score");
                throw;
            }
        }

        private async Task EnsureStudentCanBeGradedAsync(string studentId)
        {
            var student = await _students.Find(s => s.Id == studentId).FirstOrDefaultAsync();
            if (student == null)
            {
                throw new InvalidOperationException($"Không tìm thấy học sinh: {studentId}");
            }

            var normalizedStatus = (student.Status ?? string.Empty).Trim().ToLowerInvariant();
            var isInactiveByStatus = normalizedStatus == "inactive";
            if (!student.IsActive || isInactiveByStatus)
            {
                throw new InvalidOperationException(
                    $"Học sinh '{student.MiddleName} {student.FirstName}' đang ở trạng thái ngừng, không thể chấm điểm.");
            }
        }

        public async Task<List<Score>> BulkCreateOrUpdateScoresAsync(BulkScoreRequest request, string gradedBy)
        {
            var results = new List<Score>();

            try
            {
                ValidateBulkRequest(request);

                var dedupedItems = request.Scores
                    .Where(item => item != null)
                    .GroupBy(item => item.StudentId.Trim(), StringComparer.Ordinal)
                    .Select(group => group.Last())
                    .ToList();

                foreach (var item in dedupedItems)
                {
                    var scoreRequest = new CreateScoreRequest
                    {
                        StudentId = item.StudentId.Trim(),
                        AssignmentId = request.AssignmentId.Trim(),
                        ClassId = request.ClassId.Trim(),
                        ScoreValue = item.ScoreValue,
                        Feedback = item.Feedback,
                        AutoGradingErrors = item.AutoGradingErrors,
                        AutoGradingTaskResults = item.AutoGradingTaskResults
                    };

                    var score = await CreateOrUpdateScoreAsync(scoreRequest, gradedBy);
                    results.Add(score);
                }

                _logger.LogInformation("✅ Bulk scores updated: {Count} scores by {GradedBy}",
                    results.Count, gradedBy);

                return results;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (FormatException ex)
            {
                _logger.LogWarning(ex, "⚠️ Invalid bulk score payload format");
                throw new InvalidOperationException("Dữ liệu lưu điểm hàng loạt không hợp lệ.");
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(ex, "⚠️ Invalid bson payload when bulk saving scores");
                throw new InvalidOperationException("Dữ liệu lưu điểm hàng loạt không hợp lệ.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error bulk updating scores");
                throw;
            }
        }

        public async Task<bool> DeleteScoreAsync(string id, string userId)
        {
            try
            {
                var result = await _scores.DeleteOneAsync(s => s.Id == id);

                if (result.DeletedCount > 0)
                {
                    _logger.LogInformation("✅ Score deleted: {Id} by user {UserId}", id, userId);
                    return true;
                }

                _logger.LogWarning("⚠️ Score not found: {Id}", id);
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error deleting score {Id}", id);
                throw;
            }
        }

        public async Task<List<ScoreResponse>> GetScoresByClassAsync(string classId)
        {
            try
            {
                EnsureValidObjectId(classId, "Mã lớp");

                var scores = await _scores.Find(s => s.ClassId == classId).ToListAsync();
                return await BuildScoreResponsesForClassAsync(classId, scores);
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(
                    ex,
                    "⚠️ Deserialization error while reading scores for class {ClassId}. Falling back to tolerant document mapping.",
                    classId);

                return await GetScoresByClassFromRawDocumentsAsync(classId);
            }
            catch (FormatException ex)
            {
                _logger.LogWarning(
                    ex,
                    "⚠️ Invalid score document format while reading scores for class {ClassId}. Falling back to tolerant document mapping.",
                    classId);

                return await GetScoresByClassFromRawDocumentsAsync(classId);
            }
            catch (Exception ex)
            {
                var diagnostic = await BuildClassScoreSchemaDiagnosticAsync(classId);
                _logger.LogError(
                    ex,
                    "❌ Error getting all scores for class {ClassId}. Diagnostic: {Diagnostic}",
                    classId,
                    diagnostic);
                throw new InvalidOperationException(
                    $"Không thể đọc điểm của lớp. {diagnostic}",
                    ex);
            }
        }

        private void ValidateBulkRequest(BulkScoreRequest request)
        {
            if (request == null)
            {
                throw new InvalidOperationException("Thiếu dữ liệu lưu điểm hàng loạt.");
            }

            EnsureValidObjectId(request.AssignmentId, "Mã bài tập");
            EnsureValidObjectId(request.ClassId, "Mã lớp");

            if (request.Scores == null || request.Scores.Count == 0)
            {
                throw new InvalidOperationException("Danh sách điểm trống, không có dữ liệu để lưu.");
            }

            foreach (var item in request.Scores)
            {
                if (item == null)
                {
                    throw new InvalidOperationException("Có bản ghi điểm không hợp lệ trong danh sách.");
                }

                EnsureValidObjectId(item.StudentId, "Mã học sinh");
                _ = NormalizeScoreValue(item.ScoreValue);
            }
        }

        private void ValidateCreateOrUpdateRequest(CreateScoreRequest request)
        {
            if (request == null)
            {
                throw new InvalidOperationException("Thiếu dữ liệu lưu điểm.");
            }

            EnsureValidObjectId(request.StudentId, "Mã học sinh");
            EnsureValidObjectId(request.AssignmentId, "Mã bài tập");
            EnsureValidObjectId(request.ClassId, "Mã lớp");
            _ = NormalizeScoreValue(request.ScoreValue);
        }

        private static void EnsureValidObjectId(string? value, string fieldLabel)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new InvalidOperationException($"{fieldLabel} không hợp lệ.");
            }
        }

        private static double? NormalizeScoreValue(double? rawValue)
        {
            if (!rawValue.HasValue)
            {
                return null;
            }

            var value = rawValue.Value;
            if (!double.IsFinite(value))
            {
                throw new InvalidOperationException("Điểm không hợp lệ.");
            }

            if (value < 0)
            {
                throw new InvalidOperationException("Điểm không được nhỏ hơn 0.");
            }

            if (value > 1000)
            {
                throw new InvalidOperationException("Điểm không được vượt quá 1000.");
            }

            return Math.Round(value, 2, MidpointRounding.AwayFromZero);
        }

        private static string? NormalizeFeedback(string? feedback)
        {
            if (string.IsNullOrWhiteSpace(feedback))
            {
                return null;
            }

            var trimmed = feedback.Trim();
            return trimmed.Length <= 500 ? trimmed : trimmed[..500];
        }

        private static List<string> NormalizeAutoGradingErrors(List<string>? autoGradingErrors)
        {
            if (autoGradingErrors == null || autoGradingErrors.Count == 0)
            {
                return new List<string>();
            }

            return autoGradingErrors
                .Where(error => !string.IsNullOrWhiteSpace(error))
                .Select(error => error.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(100)
                .ToList();
        }

        private static List<string> NormalizeSingleTaskFeedbackLine(List<string>? lines)
        {
            var normalizedLines = NormalizeAutoGradingErrors(lines);
            if (normalizedLines.Count <= 1)
            {
                return normalizedLines;
            }

            return new List<string> { normalizedLines[0] };
        }

        private static List<AutoGradingTaskResultRequest> NormalizeAutoGradingTaskResults(
            List<AutoGradingTaskResultRequest>? autoGradingTaskResults)
        {
            if (autoGradingTaskResults == null || autoGradingTaskResults.Count == 0)
            {
                return new List<AutoGradingTaskResultRequest>();
            }

            var normalized = new List<AutoGradingTaskResultRequest>();
            for (var index = 0; index < autoGradingTaskResults.Count; index++)
            {
                var item = autoGradingTaskResults[index];
                if (item == null)
                {
                    continue;
                }

                var taskId = (item.TaskId ?? string.Empty).Trim();
                var taskName = (item.TaskName ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(taskId))
                {
                    taskId = $"TASK-{index + 1:00}";
                }

                if (string.IsNullOrWhiteSpace(taskName))
                {
                    taskName = taskId;
                }

                var normalizedMaxScore = double.IsFinite(item.MaxScore) && item.MaxScore > 0
                    ? Math.Round(item.MaxScore, 4, MidpointRounding.AwayFromZero)
                    : 1d;

                var normalizedScore = double.IsFinite(item.Score)
                    ? Math.Round(Math.Max(0d, item.Score), 4, MidpointRounding.AwayFromZero)
                    : 0d;

                if (item.IsPassed && normalizedScore < normalizedMaxScore * 0.5d)
                {
                    normalizedScore = normalizedMaxScore;
                }
                else if (!item.IsPassed && normalizedScore >= normalizedMaxScore * 0.5d)
                {
                    normalizedScore = 0d;
                }

                var normalizedErrors = NormalizeSingleTaskFeedbackLine(item.Errors);
                var normalizedDetails = NormalizeAutoGradingErrors(item.Details);
                var normalizedFixActions = NormalizeSingleTaskFeedbackLine(item.FixActions);

                normalized.Add(new AutoGradingTaskResultRequest
                {
                    TaskId = taskId,
                    TaskName = taskName,
                    Score = normalizedScore,
                    MaxScore = normalizedMaxScore,
                    IsPassed = item.IsPassed,
                    Details = normalizedDetails,
                    Errors = normalizedErrors,
                    FixActions = normalizedFixActions
                });
            }

            return normalized;
        }

        private List<TaskResult> BuildAnalyticsTaskResults(CreateScoreRequest request)
        {
            if (request.AutoGradingTaskResults != null && request.AutoGradingTaskResults.Count > 0)
            {
                return request.AutoGradingTaskResults
                    .Select(task => new TaskResult
                    {
                        TaskId = task.TaskId,
                        TaskName = task.TaskName,
                        Score = (decimal)task.Score,
                        MaxScore = (decimal)task.MaxScore,
                        Details = task.Details ?? new List<string>(),
                        Errors = task.Errors ?? new List<string>(),
                        FixActions = new SingleFixActionList(task.FixActions)
                    })
                    .ToList();
            }

            var normalizedAutoErrors = NormalizeAutoGradingErrors(request.AutoGradingErrors);
            if (normalizedAutoErrors.Count == 0)
            {
                return new List<TaskResult>();
            }

            var taskOrder = new List<string>();
            var taskNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var taskErrors = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var taskErrorSets = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            for (var index = 0; index < normalizedAutoErrors.Count; index++)
            {
                var rawError = normalizedAutoErrors[index];
                var taskId = $"ERROR-{index + 1:00}";
                var taskName = taskId;
                var errorMessage = rawError;

                var matched = AutoErrorQuestionRegex.Match(rawError);
                if (matched.Success)
                {
                    var question = matched.Groups["question"].Value;
                    var extractedTaskName = matched.Groups["taskName"].Value.Trim();
                    var extractedMessage = matched.Groups["message"].Value.Trim();

                    taskId = $"Cau {question}";
                    taskName = string.IsNullOrWhiteSpace(extractedTaskName) ? taskId : extractedTaskName;
                    errorMessage = string.IsNullOrWhiteSpace(extractedMessage) ? rawError : extractedMessage;
                }

                if (!taskErrors.ContainsKey(taskId))
                {
                    taskOrder.Add(taskId);
                    taskNames[taskId] = taskName;
                    taskErrors[taskId] = new List<string>();
                    taskErrorSets[taskId] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                if (taskErrorSets[taskId].Add(errorMessage))
                {
                    taskErrors[taskId].Add(errorMessage);
                }
            }

            return taskOrder
                .Select(taskId => new TaskResult
                {
                    TaskId = taskId,
                    TaskName = taskNames[taskId],
                    Score = 0m,
                    MaxScore = 1m,
                    Details = new List<string>(),
                    Errors = taskErrors[taskId]
                })
                .ToList();
        }

        private async Task<string?> ResolveGraderNameAsync(string? graderId)
        {
            if (string.IsNullOrWhiteSpace(graderId))
            {
                return null;
            }

            var normalizedGraderId = graderId.Trim();
            if (!ObjectId.TryParse(normalizedGraderId, out _))
            {
                return null;
            }

            try
            {
                var grader = await _users.Find(u => u.Id == normalizedGraderId).FirstOrDefaultAsync();
                return grader?.FullName ?? grader?.Username;
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(ex, "⚠️ Skipping invalid grader payload for user {UserId}", normalizedGraderId);
                return null;
            }
        }

        private async Task<List<ScoreResponse>> BuildScoreResponsesAsync(List<Score> scores, Assignment? assignment)
        {
            var result = new List<ScoreResponse>();

            foreach (var score in scores)
            {
                var student = await SafeGetStudentAsync(score.StudentId);
                var graderName = await ResolveGraderNameAsync(score.GradedBy);

                result.Add(new ScoreResponse
                {
                    Id = score.Id,
                    StudentId = score.StudentId,
                    StudentFirstName = student?.FirstName ?? "",
                    StudentMiddleName = student?.MiddleName ?? "",
                    StudentFullName = $"{student?.MiddleName} {student?.FirstName}".Trim(),
                    AssignmentId = score.AssignmentId,
                    AssignmentName = assignment?.Name ?? "",
                    ScoreValue = score.ScoreValue,
                    Feedback = score.Feedback,
                    AutoGradingErrors = score.AutoGradingErrors ?? new List<string>(),
                    GradedAt = score.GradedAt,
                    GradedBy = score.GradedBy,
                    GradedByName = graderName
                });
            }

            return result.OrderBy(r => r.StudentFullName).ToList();
        }

        private async Task<List<ScoreResponse>> BuildScoreResponsesForClassAsync(string classId, List<Score> scores)
        {
            var result = new List<ScoreResponse>();

            var assignments = await SafeGetAssignmentsByClassAsync(classId);
            var assignmentsById = assignments.ToDictionary(assignment => assignment.Id, StringComparer.Ordinal);
            var students = await SafeGetStudentsByClassAsync(classId);
            var studentsById = students
                .Where(student => !string.IsNullOrWhiteSpace(student.Id))
                .ToDictionary(student => student.Id!, StringComparer.Ordinal);

            foreach (var score in scores)
            {
                studentsById.TryGetValue(score.StudentId, out var student);
                assignmentsById.TryGetValue(score.AssignmentId, out var assignment);
                var graderName = await ResolveGraderNameAsync(score.GradedBy);

                result.Add(new ScoreResponse
                {
                    Id = score.Id,
                    StudentId = score.StudentId,
                    StudentFirstName = student?.FirstName ?? "",
                    StudentMiddleName = student?.MiddleName ?? "",
                    StudentFullName = $"{student?.MiddleName} {student?.FirstName}".Trim(),
                    AssignmentId = score.AssignmentId,
                    AssignmentName = assignment?.Name ?? "",
                    ScoreValue = score.ScoreValue,
                    Feedback = score.Feedback,
                    AutoGradingErrors = score.AutoGradingErrors ?? new List<string>(),
                    GradedAt = score.GradedAt,
                    GradedBy = score.GradedBy,
                    GradedByName = graderName
                });
            }

            return result.OrderBy(r => r.StudentFullName).ToList();
        }

        private async Task<List<ScoreResponse>> GetScoresByAssignmentFromRawDocumentsAsync(string assignmentId)
        {
            var assignment = await SafeGetAssignmentAsync(assignmentId);
            var assignmentFilter = Builders<BsonDocument>.Filter.Eq("assignmentId", assignmentId);
            if (ObjectId.TryParse(assignmentId, out var assignmentObjectId))
            {
                assignmentFilter |= Builders<BsonDocument>.Filter.Eq("assignmentId", assignmentObjectId);
            }

            var documents = await _scoreDocuments.Find(assignmentFilter).ToListAsync();
            var result = new List<ScoreResponse>();

            foreach (var document in documents)
            {
                if (!TryMapScoreDocument(document, out var score))
                {
                    var documentId = ReadString(document, "_id") ?? "(unknown)";
                    _logger.LogWarning(
                        "⚠️ Skipping malformed score document {ScoreId} for assignment {AssignmentId}",
                        documentId,
                        assignmentId);
                    continue;
                }

                var student = await SafeGetStudentAsync(score.StudentId);
                var graderName = await ResolveGraderNameAsync(score.GradedBy);

                result.Add(new ScoreResponse
                {
                    Id = score.Id,
                    StudentId = score.StudentId,
                    StudentFirstName = student?.FirstName ?? "",
                    StudentMiddleName = student?.MiddleName ?? "",
                    StudentFullName = $"{student?.MiddleName} {student?.FirstName}".Trim(),
                    AssignmentId = score.AssignmentId,
                    AssignmentName = assignment?.Name ?? "",
                    ScoreValue = score.ScoreValue,
                    Feedback = score.Feedback,
                    AutoGradingErrors = score.AutoGradingErrors ?? new List<string>(),
                    GradedAt = score.GradedAt,
                    GradedBy = score.GradedBy,
                    GradedByName = graderName
                });
            }

            return result.OrderBy(r => r.StudentFullName).ToList();
        }

        private async Task<List<ScoreResponse>> GetScoresByClassFromRawDocumentsAsync(string classId)
        {
            var classFilter = Builders<BsonDocument>.Filter.Eq("classId", classId);
            if (ObjectId.TryParse(classId, out var classObjectId))
            {
                classFilter |= Builders<BsonDocument>.Filter.Eq("classId", classObjectId);
            }

            var documents = await _scoreDocuments.Find(classFilter).ToListAsync();
            var assignments = await SafeGetAssignmentsByClassAsync(classId);
            var assignmentsById = assignments.ToDictionary(assignment => assignment.Id, StringComparer.Ordinal);
            var students = await SafeGetStudentsByClassAsync(classId);
            var studentsById = students
                .Where(student => !string.IsNullOrWhiteSpace(student.Id))
                .ToDictionary(student => student.Id!, StringComparer.Ordinal);
            var result = new List<ScoreResponse>();

            foreach (var document in documents)
            {
                if (!TryMapScoreDocument(document, out var score))
                {
                    var documentId = ReadString(document, "_id") ?? "(unknown)";
                    _logger.LogWarning(
                        "⚠️ Skipping malformed score document {ScoreId} for class {ClassId}",
                        documentId,
                        classId);
                    continue;
                }

                studentsById.TryGetValue(score.StudentId, out var student);
                assignmentsById.TryGetValue(score.AssignmentId, out var assignment);
                var graderName = await ResolveGraderNameAsync(score.GradedBy);

                result.Add(new ScoreResponse
                {
                    Id = score.Id,
                    StudentId = score.StudentId,
                    StudentFirstName = student?.FirstName ?? "",
                    StudentMiddleName = student?.MiddleName ?? "",
                    StudentFullName = $"{student?.MiddleName} {student?.FirstName}".Trim(),
                    AssignmentId = score.AssignmentId,
                    AssignmentName = assignment?.Name ?? "",
                    ScoreValue = score.ScoreValue,
                    Feedback = score.Feedback,
                    AutoGradingErrors = score.AutoGradingErrors ?? new List<string>(),
                    GradedAt = score.GradedAt,
                    GradedBy = score.GradedBy,
                    GradedByName = graderName
                });
            }

            return result.OrderBy(r => r.StudentFullName).ToList();
        }

        private async Task<Student?> SafeGetStudentAsync(string? studentId)
        {
            if (string.IsNullOrWhiteSpace(studentId) || !ObjectId.TryParse(studentId.Trim(), out _))
            {
                return null;
            }

            try
            {
                return await _students.Find(s => s.Id == studentId).FirstOrDefaultAsync();
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(ex, "⚠️ Skipping invalid student payload for student {StudentId}", studentId);
                return null;
            }
        }

        private async Task<List<Student>> SafeGetStudentsByClassAsync(string classId)
        {
            try
            {
                return await _students.Find(s => s.ClassId == classId).ToListAsync();
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(ex, "⚠️ Unable to fully deserialize students for class {ClassId}; returning partial empty list.", classId);
                return new List<Student>();
            }
        }

        private async Task<List<Assignment>> SafeGetAssignmentsByClassAsync(string classId)
        {
            try
            {
                return await _assignments.Find(a => a.ClassId == classId).ToListAsync();
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(ex, "⚠️ Unable to fully deserialize assignments for class {ClassId}; returning partial empty list.", classId);
                return new List<Assignment>();
            }
        }

        private async Task<string> BuildScoreSchemaDiagnosticAsync(string assignmentId)
        {
            try
            {
                var assignment = await _assignmentDocuments
                    .Find(Builders<BsonDocument>.Filter.Eq("_id", ObjectId.Parse(assignmentId)))
                    .FirstOrDefaultAsync();

                var endpoint = assignment == null ? "(không rõ endpoint)" : ReadString(assignment, "gradingApiEndpoint") ?? "(trống)";
                var assignmentName = assignment == null ? "(không rõ tên)" : ReadString(assignment, "name") ?? "(trống)";

                var assignmentFilter = Builders<BsonDocument>.Filter.Eq("assignmentId", assignmentId);
                if (ObjectId.TryParse(assignmentId, out var assignmentObjectId))
                {
                    assignmentFilter |= Builders<BsonDocument>.Filter.Eq("assignmentId", assignmentObjectId);
                }

                var rawScores = await _scoreDocuments.Find(assignmentFilter).Limit(5).ToListAsync();
                if (rawScores.Count == 0)
                {
                    return $"Assignment '{assignmentName}' ({endpoint}) chưa có score document nào.";
                }

                foreach (var score in rawScores)
                {
                    var problems = InspectScoreDocumentShape(score);
                    if (problems.Count > 0)
                    {
                        var scoreId = ReadString(score, "_id") ?? "(unknown)";
                        return $"Assignment '{assignmentName}' ({endpoint}) có score {scoreId} lệch schema: {string.Join("; ", problems)}";
                    }
                }

                var fieldSummary = SummarizeScoreFieldTypes(rawScores);
                return $"Assignment '{assignmentName}' ({endpoint}) không có field bắt buộc sai rõ ràng, nhưng schema thực tế là: {fieldSummary}";
            }
            catch (Exception diagnosticEx)
            {
                _logger.LogWarning(diagnosticEx, "⚠️ Unable to build score schema diagnostic for assignment {AssignmentId}", assignmentId);
                return "Không lấy được chẩn đoán schema từ raw score documents.";
            }
        }

        private async Task<string> BuildClassScoreSchemaDiagnosticAsync(string classId)
        {
            try
            {
                var classFilter = Builders<BsonDocument>.Filter.Eq("classId", classId);
                if (ObjectId.TryParse(classId, out var classObjectId))
                {
                    classFilter |= Builders<BsonDocument>.Filter.Eq("classId", classObjectId);
                }

                var rawScores = await _scoreDocuments.Find(classFilter).Limit(5).ToListAsync();
                if (rawScores.Count == 0)
                {
                    return $"Lớp {classId} chưa có score document nào hoặc score đang lưu bằng schema/filter khác.";
                }

                foreach (var score in rawScores)
                {
                    var problems = InspectScoreDocumentShape(score);
                    if (problems.Count > 0)
                    {
                        var scoreId = ReadString(score, "_id") ?? "(unknown)";
                        return $"Lớp {classId} có score {scoreId} lệch schema: {string.Join("; ", problems)}";
                    }
                }

                var fieldSummary = SummarizeScoreFieldTypes(rawScores);
                return $"Schema điểm của lớp {classId}: {fieldSummary}";
            }
            catch (Exception diagnosticEx)
            {
                _logger.LogWarning(diagnosticEx, "⚠️ Unable to build class score schema diagnostic for class {ClassId}", classId);
                return "Không lấy được chẩn đoán schema từ raw class score documents.";
            }
        }

        private static List<string> InspectScoreDocumentShape(BsonDocument score)
        {
            var problems = new List<string>();

            ValidateObjectIdLikeField(score, "_id", problems);
            ValidateObjectIdLikeField(score, "studentId", problems);
            ValidateObjectIdLikeField(score, "assignmentId", problems);
            ValidateObjectIdLikeField(score, "classId", problems);
            ValidateOptionalObjectIdLikeField(score, "gradedBy", problems);
            ValidateOptionalObjectIdLikeField(score, "createdBy", problems);
            ValidateOptionalObjectIdLikeField(score, "updatedBy", problems);
            ValidateOptionalNumericField(score, "scoreValue", problems);
            ValidateOptionalStringField(score, "feedback", problems);
            ValidateOptionalStringArrayField(score, "autoGradingErrors", problems);
            ValidateOptionalDateField(score, "gradedAt", problems);
            ValidateOptionalDateField(score, "createdAt", problems);
            ValidateOptionalDateField(score, "updatedAt", problems);

            return problems;
        }

        private static void ValidateObjectIdLikeField(BsonDocument document, string fieldName, List<string> problems)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                problems.Add($"{fieldName} bị thiếu");
                return;
            }

            if (value.BsonType == BsonType.ObjectId)
            {
                return;
            }

            if (value.BsonType == BsonType.String && ObjectId.TryParse(value.AsString, out _))
            {
                return;
            }

            problems.Add($"{fieldName} có kiểu {DescribeBsonValue(value)}");
        }

        private static void ValidateOptionalObjectIdLikeField(BsonDocument document, string fieldName, List<string> problems)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return;
            }

            if (value.BsonType == BsonType.ObjectId)
            {
                return;
            }

            if (value.BsonType == BsonType.String && ObjectId.TryParse(value.AsString, out _))
            {
                return;
            }

            problems.Add($"{fieldName} có kiểu {DescribeBsonValue(value)}");
        }

        private static void ValidateOptionalNumericField(BsonDocument document, string fieldName, List<string> problems)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return;
            }

            if (value.BsonType is BsonType.Double or BsonType.Int32 or BsonType.Int64 or BsonType.Decimal128)
            {
                return;
            }

            if (value.BsonType == BsonType.String && double.TryParse(value.AsString, out _))
            {
                return;
            }

            problems.Add($"{fieldName} có kiểu {DescribeBsonValue(value)}");
        }

        private static void ValidateOptionalStringField(BsonDocument document, string fieldName, List<string> problems)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return;
            }

            if (value.BsonType != BsonType.String)
            {
                problems.Add($"{fieldName} có kiểu {DescribeBsonValue(value)}");
            }
        }

        private static void ValidateOptionalStringArrayField(BsonDocument document, string fieldName, List<string> problems)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return;
            }

            if (value.BsonType != BsonType.Array)
            {
                problems.Add($"{fieldName} có kiểu {DescribeBsonValue(value)}");
                return;
            }

            var invalidElement = value.AsBsonArray.FirstOrDefault(item => item is { IsBsonNull: false, BsonType: not BsonType.String });
            if (invalidElement != null)
            {
                problems.Add($"{fieldName} chứa phần tử kiểu {DescribeBsonValue(invalidElement)}");
            }
        }

        private static void ValidateOptionalDateField(BsonDocument document, string fieldName, List<string> problems)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return;
            }

            if (value.BsonType == BsonType.DateTime)
            {
                return;
            }

            if (value.BsonType == BsonType.String && DateTime.TryParse(value.AsString, out _))
            {
                return;
            }

            problems.Add($"{fieldName} có kiểu {DescribeBsonValue(value)}");
        }

        private static string SummarizeScoreFieldTypes(List<BsonDocument> rawScores)
        {
            var fields = new[]
            {
                "_id",
                "studentId",
                "assignmentId",
                "classId",
                "scoreValue",
                "feedback",
                "autoGradingErrors",
                "gradedAt",
                "gradedBy",
                "createdAt",
                "createdBy",
                "updatedAt",
                "updatedBy"
            };

            var summaries = new List<string>();
            foreach (var field in fields)
            {
                var types = rawScores
                    .Where(doc => doc.Contains(field))
                    .Select(doc => DescribeBsonValue(doc[field]))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (types.Count == 0)
                {
                    continue;
                }

                summaries.Add($"{field}=[{string.Join(", ", types)}]");
            }

            return string.Join("; ", summaries);
        }

        private static string DescribeBsonValue(BsonValue value)
        {
            if (value.IsBsonNull)
            {
                return "Null";
            }

            if (value.BsonType != BsonType.Array)
            {
                return value.BsonType.ToString();
            }

            var elementTypes = value.AsBsonArray
                .Select(item => item.IsBsonNull ? "Null" : item.BsonType.ToString())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return elementTypes.Count == 0
                ? "Array<empty>"
                : $"Array<{string.Join("|", elementTypes)}>";
        }

        private async Task<Assignment?> SafeGetAssignmentAsync(string? assignmentId)
        {
            if (string.IsNullOrWhiteSpace(assignmentId) || !ObjectId.TryParse(assignmentId.Trim(), out _))
            {
                return null;
            }

            try
            {
                return await _assignments.Find(a => a.Id == assignmentId).FirstOrDefaultAsync();
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(ex, "⚠️ Skipping invalid assignment payload for assignment {AssignmentId}", assignmentId);
                return null;
            }
        }

        private static bool TryMapScoreDocument(BsonDocument document, out Score score)
        {
            score = new Score();

            try
            {
                var id = ReadString(document, "_id");
                var studentId = ReadString(document, "studentId");
                var assignmentId = ReadString(document, "assignmentId");
                var classId = ReadString(document, "classId");

                if (string.IsNullOrWhiteSpace(id) ||
                    string.IsNullOrWhiteSpace(studentId) ||
                    string.IsNullOrWhiteSpace(assignmentId) ||
                    string.IsNullOrWhiteSpace(classId))
                {
                    return false;
                }

                score = new Score
                {
                    Id = id,
                    StudentId = studentId,
                    AssignmentId = assignmentId,
                    ClassId = classId,
                    ScoreValue = ReadDouble(document, "scoreValue"),
                    Feedback = ReadString(document, "feedback"),
                    AutoGradingErrors = ReadStringList(document, "autoGradingErrors"),
                    GradedAt = ReadDateTime(document, "gradedAt"),
                    GradedBy = ReadString(document, "gradedBy"),
                    CreatedAt = ReadDateTime(document, "createdAt") ?? DateTime.UtcNow,
                    CreatedBy = ReadString(document, "createdBy"),
                    UpdatedAt = ReadDateTime(document, "updatedAt"),
                    UpdatedBy = ReadString(document, "updatedBy")
                };

                return true;
            }
            catch
            {
                score = new Score();
                return false;
            }
        }

        private static string? ReadString(BsonDocument document, string fieldName)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return null;
            }

            return value.BsonType switch
            {
                BsonType.ObjectId => value.AsObjectId.ToString(),
                BsonType.String => value.AsString,
                _ => value.ToString()
            };
        }

        private static double? ReadDouble(BsonDocument document, string fieldName)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return null;
            }

            return value.BsonType switch
            {
                BsonType.Double => value.AsDouble,
                BsonType.Int32 => value.AsInt32,
                BsonType.Int64 => value.AsInt64,
                BsonType.Decimal128 => (double)value.AsDecimal128,
                BsonType.String when double.TryParse(value.AsString, out var parsed) => parsed,
                _ => null
            };
        }

        private static DateTime? ReadDateTime(BsonDocument document, string fieldName)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return null;
            }

            return value.BsonType switch
            {
                BsonType.DateTime => value.ToUniversalTime(),
                BsonType.String when DateTime.TryParse(value.AsString, out var parsed) => parsed,
                _ => null
            };
        }

        private static List<string> ReadStringList(BsonDocument document, string fieldName)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return new List<string>();
            }

            if (value.BsonType != BsonType.Array)
            {
                var singleValue = ReadString(document, fieldName);
                return string.IsNullOrWhiteSpace(singleValue) ? new List<string>() : new List<string> { singleValue };
            }

            return value.AsBsonArray
                .Select(item => item.IsBsonNull ? null : item.ToString())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(item => item!.Trim())
                .ToList();
        }

        private async Task SaveGradingAttemptFromSavedScoreAsync(
            CreateScoreRequest request,
            Assignment? assignment,
            string gradedBy)
        {
            try
            {
                var normalizedEndpoint = GradingApiEndpoints.NormalizeEndpoint(assignment?.GradingApiEndpoint);
                var projectEndpoint = string.IsNullOrWhiteSpace(normalizedEndpoint)
                    ? GradingApiEndpoints.Project09
                    : normalizedEndpoint;

                var projectId = "MANUAL";
                if (GradingApiEndpoints.TryExtractProjectNumber(projectEndpoint, out var projectNumber))
                {
                    projectId = $"P{projectNumber:00}";
                }

                var maxScoreValue = assignment?.MaxScore ?? 0d;
                if (maxScoreValue <= 0d)
                {
                    maxScoreValue = 100d;
                }

                var scoreValue = request.ScoreValue ?? 0d;
                var taskResults = BuildAnalyticsTaskResults(request);

                var gradingResult = new GradingResult
                {
                    ProjectId = projectId,
                    ProjectName = assignment?.Name ?? "Saved Score",
                    TotalScore = (decimal)scoreValue,
                    MaxScore = (decimal)maxScoreValue,
                    TaskResults = taskResults
                };

                await _analyticsService.SaveGradingAttemptAsync(
                    gradingResult,
                    projectEndpoint,
                    request.ClassId,
                    request.AssignmentId,
                    request.StudentId,
                    gradedBy,
                    persistToDatabase: true);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "⚠️ Không thể lưu lịch sử lần chấm khi lưu điểm học sinh");
            }
        }
    }
}

