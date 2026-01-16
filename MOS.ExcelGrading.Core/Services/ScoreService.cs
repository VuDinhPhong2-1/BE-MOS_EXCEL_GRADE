// MOS.ExcelGrading.Core/Services/ScoreService.cs
using Microsoft.Extensions.Logging;
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
        private readonly ILogger<ScoreService> _logger;

        public ScoreService(
            IMongoDatabase database,
            ILogger<ScoreService> logger)
        {
            _scores = database.GetCollection<Score>("scores");
            _students = database.GetCollection<Student>("students");
            _assignments = database.GetCollection<Assignment>("assignments");
            _users = database.GetCollection<User>("users");
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

                    string? graderName = null;
                    if (!string.IsNullOrEmpty(score.GradedBy))
                    {
                        var grader = await _users.Find(u => u.Id == score.GradedBy).FirstOrDefaultAsync();
                        graderName = grader?.FullName ?? grader?.Username;
                    }

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

                    string? graderName = null;
                    if (!string.IsNullOrEmpty(score.GradedBy))
                    {
                        var grader = await _users.Find(u => u.Id == score.GradedBy).FirstOrDefaultAsync();
                        graderName = grader?.FullName ?? grader?.Username;
                    }

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
                    throw new Exception($"Student {studentId} not found");

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
                var existingScore = await GetScoreAsync(request.StudentId, request.AssignmentId);

                if (existingScore != null)
                {
                    // Update existing score
                    var update = Builders<Score>.Update
                        .Set(s => s.ScoreValue, request.ScoreValue)
                        .Set(s => s.Feedback, request.Feedback)
                        .Set(s => s.GradedAt, DateTime.UtcNow)
                        .Set(s => s.GradedBy, gradedBy)
                        .Set(s => s.UpdatedAt, DateTime.UtcNow)
                        .Set(s => s.UpdatedBy, gradedBy);

                    await _scores.UpdateOneAsync(s => s.Id == existingScore.Id, update);

                    existingScore.ScoreValue = request.ScoreValue;
                    existingScore.Feedback = request.Feedback;
                    existingScore.GradedAt = DateTime.UtcNow;
                    existingScore.GradedBy = gradedBy;

                    _logger.LogInformation("✅ Score updated for student {StudentId} by {GradedBy}",
                        request.StudentId, gradedBy);

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
                        ScoreValue = request.ScoreValue,
                        Feedback = request.Feedback,
                        GradedAt = DateTime.UtcNow,
                        GradedBy = gradedBy,
                        CreatedBy = gradedBy,
                        CreatedAt = DateTime.UtcNow
                    };

                    await _scores.InsertOneAsync(score);

                    _logger.LogInformation("✅ Score created for student {StudentId} by {GradedBy}",
                        request.StudentId, gradedBy);

                    return score;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error creating/updating score");
                throw;
            }
        }

        public async Task<List<Score>> BulkCreateOrUpdateScoresAsync(BulkScoreRequest request, string gradedBy)
        {
            var results = new List<Score>();

            try
            {
                foreach (var item in request.Scores)
                {
                    var scoreRequest = new CreateScoreRequest
                    {
                        StudentId = item.StudentId,
                        AssignmentId = request.AssignmentId,
                        ClassId = request.ClassId,
                        ScoreValue = item.ScoreValue,
                        Feedback = item.Feedback
                    };

                    var score = await CreateOrUpdateScoreAsync(scoreRequest, gradedBy);
                    results.Add(score);
                }

                _logger.LogInformation("✅ Bulk scores updated: {Count} scores by {GradedBy}",
                    results.Count, gradedBy);

                return results;
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
    }
}
