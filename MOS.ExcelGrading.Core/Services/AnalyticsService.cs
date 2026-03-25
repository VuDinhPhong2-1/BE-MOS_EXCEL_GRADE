using Microsoft.Extensions.Logging;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class AnalyticsService : IAnalyticsService
    {
        private readonly IMongoCollection<GradingAttempt> _gradingAttempts;
        private readonly ILogger<AnalyticsService> _logger;

        public AnalyticsService(IMongoDatabase database, ILogger<AnalyticsService> logger)
        {
            _gradingAttempts = database.GetCollection<GradingAttempt>("gradingAttempts");
            _logger = logger;
        }

        public async Task SaveGradingAttemptAsync(
            GradingResult result,
            string projectEndpoint,
            string? classId,
            string? assignmentId,
            string? studentId,
            string? gradedBy,
            bool persistToDatabase = false)
        {
            if (!persistToDatabase)
            {
                return;
            }

            try
            {
                var attempt = new GradingAttempt
                {
                    ProjectEndpoint = projectEndpoint,
                    ProjectId = result.ProjectId,
                    ClassId = classId,
                    AssignmentId = assignmentId,
                    StudentId = studentId,
                    TotalScore = result.TotalScore,
                    MaxScore = result.MaxScore,
                    Percentage = result.Percentage,
                    Status = result.Status,
                    GradedBy = gradedBy,
                    GradedAt = result.GradedAt,
                    TaskResults = result.TaskResults.Select(t => new GradingAttemptTask
                    {
                        TaskId = t.TaskId,
                        TaskName = t.TaskName,
                        Score = t.Score,
                        MaxScore = t.MaxScore,
                        IsPassed = t.IsPassed,
                        ErrorCount = t.Errors.Count
                    }).ToList()
                };

                await _gradingAttempts.InsertOneAsync(attempt);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving grading attempt analytics");
            }
        }

        public async Task<ClassAnalyticsOverviewResponse> GetClassOverviewAsync(string classId)
        {
            var attempts = await _gradingAttempts.Find(a => a.ClassId == classId).ToListAsync();

            var totalAttempts = attempts.Count;
            var totalStudents = attempts
                .Where(a => !string.IsNullOrWhiteSpace(a.StudentId))
                .Select(a => a.StudentId!)
                .Distinct()
                .Count();

            var averagePercentage = totalAttempts > 0 ? Math.Round(attempts.Average(a => a.Percentage), 2) : 0;
            var passRate = totalAttempts > 0
                ? Math.Round((double)attempts.Count(a => a.Percentage >= 60) / totalAttempts * 100, 2)
                : 0;
            var warningRate = totalAttempts > 0
                ? Math.Round((double)attempts.Count(a => a.Percentage < 40) / totalAttempts * 100, 2)
                : 0;

            return new ClassAnalyticsOverviewResponse
            {
                ClassId = classId,
                TotalAttempts = totalAttempts,
                TotalStudents = totalStudents,
                AveragePercentage = averagePercentage,
                PassRate = passRate,
                WarningRate = warningRate
            };
        }

        public async Task<List<WeakTaskResponse>> GetWeakTasksAsync(string classId, string? projectEndpoint, int top)
        {
            var filter = Builders<GradingAttempt>.Filter.Eq(a => a.ClassId, classId);
            if (!string.IsNullOrWhiteSpace(projectEndpoint))
            {
                filter &= Builders<GradingAttempt>.Filter.Eq(a => a.ProjectEndpoint, projectEndpoint);
            }

            var attempts = await _gradingAttempts.Find(filter).ToListAsync();

            var taskStats = attempts
                .SelectMany(a => a.TaskResults)
                .GroupBy(t => new { t.TaskId, t.TaskName })
                .Select(g =>
                {
                    var attemptCount = g.Count();
                    var failedCount = g.Count(t => !t.IsPassed);
                    var failedRate = attemptCount > 0
                        ? Math.Round((double)failedCount / attemptCount * 100, 2)
                        : 0;

                    return new WeakTaskResponse
                    {
                        TaskId = g.Key.TaskId,
                        TaskName = g.Key.TaskName,
                        AttemptCount = attemptCount,
                        FailedCount = failedCount,
                        FailedRate = failedRate
                    };
                })
                .OrderByDescending(t => t.FailedRate)
                .ThenByDescending(t => t.FailedCount)
                .Take(Math.Max(1, top))
                .ToList();

            return taskStats;
        }

        public async Task<List<ProjectPerformanceResponse>> GetProjectPerformanceAsync(string classId)
        {
            var attempts = await _gradingAttempts.Find(a => a.ClassId == classId).ToListAsync();

            return attempts
                .GroupBy(a => a.ProjectEndpoint)
                .Select(g =>
                {
                    var attemptCount = g.Count();
                    var passRate = attemptCount > 0
                        ? Math.Round((double)g.Count(a => a.Percentage >= 60) / attemptCount * 100, 2)
                        : 0;

                    return new ProjectPerformanceResponse
                    {
                        ProjectEndpoint = g.Key,
                        AttemptCount = attemptCount,
                        AveragePercentage = Math.Round(g.Average(a => a.Percentage), 2),
                        PassRate = passRate
                    };
                })
                .OrderBy(p => p.ProjectEndpoint)
                .ToList();
        }
    }
}
