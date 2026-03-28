// MOS.ExcelGrading.Core/Services/ScoreService.cs
using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class ScoreService : IScoreService
    {
        private readonly IMongoCollection<Score> _scores;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<Assignment> _assignments;
        private readonly IMongoCollection<User> _users;
        private readonly IAnalyticsService _analyticsService;
        private readonly ILogger<ScoreService> _logger;

        public ScoreService(
            IMongoDatabase database,
            IAnalyticsService analyticsService,
            ILogger<ScoreService> logger)
        {
            _scores = database.GetCollection<Score>("scores");
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
                var scores = await _scores.Find(s => s.AssignmentId == assignmentId).ToListAsync();
                var result = new List<ScoreResponse>();

                var assignment = await _assignments.Find(a => a.Id == assignmentId).FirstOrDefaultAsync();

                foreach (var score in scores)
                {
                    var student = await _students.Find(s => s.Id == score.StudentId).FirstOrDefaultAsync();

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
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting scores for assignment {AssignmentId}", assignmentId);
                throw;
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
                    AutoGradingErrors = normalizedAutoErrors
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
                        AutoGradingErrors = item.AutoGradingErrors
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
                var scores = await _scores.Find(s => s.ClassId == classId).ToListAsync();
                var result = new List<ScoreResponse>();

                // Lấy dữ liệu cần thiết
                var assignments = await _assignments.Find(a => a.ClassId == classId).ToListAsync();
                var students = await _students.Find(s => s.ClassId == classId).ToListAsync();

                foreach (var score in scores)
                {
                    var student = students.FirstOrDefault(st => st.Id == score.StudentId);
                    var assignment = assignments.FirstOrDefault(a => a.Id == score.AssignmentId);

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

                // Đảm bảo trả về đủ dữ liệu điểm các assignment, nếu học sinh chưa có score vẫn có thể default 0/undefined tại FE khi map
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting all scores for class {ClassId}", classId);
                throw;
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

            var grader = await _users.Find(u => u.Id == normalizedGraderId).FirstOrDefaultAsync();
            return grader?.FullName ?? grader?.Username;
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
                var autoErrors = request.AutoGradingErrors?
                    .Where(error => !string.IsNullOrWhiteSpace(error))
                    .Select(error => error.Trim())
                    .ToList() ?? new List<string>();

                var gradingResult = new GradingResult
                {
                    ProjectId = projectId,
                    ProjectName = assignment?.Name ?? "Saved Score",
                    TotalScore = (decimal)scoreValue,
                    MaxScore = (decimal)maxScoreValue,
                    TaskResults = new List<TaskResult>
                    {
                        new()
                        {
                            TaskId = "SCORE-SAVE",
                            TaskName = "Saved score snapshot",
                            Score = (decimal)scoreValue,
                            MaxScore = (decimal)maxScoreValue,
                            Errors = autoErrors
                        }
                    }
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

