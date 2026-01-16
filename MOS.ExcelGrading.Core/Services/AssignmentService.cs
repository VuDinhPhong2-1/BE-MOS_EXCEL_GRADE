// MOS.ExcelGrading.Core/Services/AssignmentService.cs
using Microsoft.Extensions.Logging;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class AssignmentService : IAssignmentService
    {
        private readonly IMongoCollection<Assignment> _assignments;
        private readonly IMongoCollection<Score> _scores;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<User> _users;
        private readonly ILogger<AssignmentService> _logger;

        public AssignmentService(
            IMongoDatabase database,
            ILogger<AssignmentService> logger)
        {
            _assignments = database.GetCollection<Assignment>("assignments");
            _scores = database.GetCollection<Score>("scores");
            _students = database.GetCollection<Student>("students");
            _users = database.GetCollection<User>("users");
            _logger = logger;
        }

        public async Task<List<Assignment>> GetAssignmentsByClassIdAsync(string classId)
        {
            try
            {
                var filter = Builders<Assignment>.Filter.Eq(a => a.ClassId, classId) &
                             Builders<Assignment>.Filter.Eq(a => a.IsActive, true);

                return await _assignments.Find(filter)
                    .SortByDescending(a => a.CreatedAt)
                    .ToListAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting assignments for class {ClassId}", classId);
                throw;
            }
        }

        public async Task<List<AssignmentWithStatsResponse>> GetAssignmentsWithStatsByClassIdAsync(string classId)
        {
            try
            {
                var assignments = await GetAssignmentsByClassIdAsync(classId);
                var result = new List<AssignmentWithStatsResponse>();

                // Get total students in class
                var totalStudents = await _students.CountDocumentsAsync(s =>
                    s.ClassId == classId && s.IsActive);

                foreach (var assignment in assignments)
                {
                    // Get scores for this assignment
                    var scores = await _scores.Find(s => s.AssignmentId == assignment.Id)
                        .ToListAsync();

                    var gradedCount = scores.Count(s => s.ScoreValue.HasValue);
                    var avgScore = scores.Any(s => s.ScoreValue.HasValue)
                        ? scores.Where(s => s.ScoreValue.HasValue)
                                .Average(s => s.ScoreValue!.Value)
                        : 0;

                    // Get creator name
                    string? creatorName = null;
                    if (!string.IsNullOrEmpty(assignment.CreatedBy))
                    {
                        var creator = await _users.Find(u => u.Id == assignment.CreatedBy)
                            .FirstOrDefaultAsync();
                        creatorName = creator?.FullName ?? creator?.Username;
                    }

                    result.Add(new AssignmentWithStatsResponse
                    {
                        Id = assignment.Id,
                        Name = assignment.Name,
                        Description = assignment.Description,
                        ClassId = assignment.ClassId,
                        MaxScore = assignment.MaxScore,
                        CreatedAt = assignment.CreatedAt,
                        IsActive = assignment.IsActive,
                        CreatedBy = assignment.CreatedBy,
                        CreatedByName = creatorName,
                        UpdatedAt = assignment.UpdatedAt,
                        TotalStudents = (int)totalStudents,
                        GradedStudents = gradedCount,
                        AverageScore = Math.Round(avgScore, 2),
                        CompletionRate = totalStudents > 0
                            ? Math.Round((double)gradedCount / totalStudents * 100, 2)
                            : 0
                    });
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting assignments with stats for class {ClassId}", classId);
                throw;
            }
        }

        public async Task<Assignment?> GetAssignmentByIdAsync(string id)
        {
            try
            {
                return await _assignments.Find(a => a.Id == id).FirstOrDefaultAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting assignment {Id}", id);
                throw;
            }
        }

        public async Task<Assignment?> UpdateAssignmentAsync(string id, UpdateAssignmentRequest request, string userId)
        {
            try
            {
                var updateDefinitions = new List<UpdateDefinition<Assignment>>();
                var builder = Builders<Assignment>.Update;

                if (!string.IsNullOrEmpty(request.Name))
                    updateDefinitions.Add(builder.Set(a => a.Name, request.Name));

                if (request.Description != null)
                    updateDefinitions.Add(builder.Set(a => a.Description, request.Description));

                if (request.MaxScore.HasValue)
                    updateDefinitions.Add(builder.Set(a => a.MaxScore, request.MaxScore.Value));

               
                if (request.IsActive.HasValue)
                    updateDefinitions.Add(builder.Set(a => a.IsActive, request.IsActive.Value));

                updateDefinitions.Add(builder.Set(a => a.UpdatedAt, DateTime.UtcNow));
                updateDefinitions.Add(builder.Set(a => a.UpdatedBy, userId));

                var update = builder.Combine(updateDefinitions);
                var result = await _assignments.FindOneAndUpdateAsync(
                    a => a.Id == id,
                    update,
                    new FindOneAndUpdateOptions<Assignment> { ReturnDocument = ReturnDocument.After }
                );

                if (result != null)
                {
                    _logger.LogInformation("✅ Assignment updated: {Id} by user {UserId}", id, userId);
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error updating assignment {Id}", id);
                throw;
            }
        }

        public async Task<bool> DeleteAssignmentAsync(string id, string userId)
        {
            try
            {
                // Soft delete: set IsActive = false
                var update = Builders<Assignment>.Update
                    .Set(a => a.IsActive, false)
                    .Set(a => a.UpdatedAt, DateTime.UtcNow)
                    .Set(a => a.UpdatedBy, userId);

                var result = await _assignments.UpdateOneAsync(a => a.Id == id, update);

                if (result.ModifiedCount > 0)
                {
                    _logger.LogInformation("✅ Assignment soft deleted: {Id} by user {UserId}", id, userId);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error deleting assignment {Id}", id);
                throw;
            }
        }

        public async Task<bool> CanUserAccessAssignment(string assignmentId, string userId)
        {
            try
            {
                var assignment = await GetAssignmentByIdAsync(assignmentId);
                if (assignment == null) return false;

                var user = await _users.Find(u => u.Id == userId).FirstOrDefaultAsync();
                if (user == null) return false;

                // Admin can access all
                if (user.Role == UserRoles.Admin) return true;

                // Teacher can access their own assignments
                if (user.Role == UserRoles.Teacher && assignment.CreatedBy == userId)
                    return true;

                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error checking user access for assignment {AssignmentId}", assignmentId);
                return false;
            }
        }


        public async Task<Assignment> CreateAssignmentAsync(CreateAssignmentRequest request, string userId)
        {
            try
            {
                // ✅ VALIDATE GRADING API ENDPOINT
                if (request.GradingType == GradingTypes.Auto)
                {
                    if (string.IsNullOrEmpty(request.GradingApiEndpoint))
                    {
                        throw new ArgumentException("GradingApiEndpoint is required when GradingType is 'auto'");
                    }

                    if (!GradingApiEndpoints.IsValidEndpoint(request.GradingApiEndpoint))
                    {
                        throw new ArgumentException($"Invalid GradingApiEndpoint: {request.GradingApiEndpoint}");
                    }
                }

                var assignment = new Assignment
                {
                    Name = request.Name,
                    Description = request.Description,
                    ClassId = request.ClassId,
                    MaxScore = request.MaxScore,
                    GradingType = request.GradingType,
                    GradingApiEndpoint = request.GradingApiEndpoint,
                    CreatedBy = userId,
                    CreatedAt = DateTime.UtcNow
                };

                await _assignments.InsertOneAsync(assignment);

                _logger.LogInformation(
                    "✅ Assignment created: {Name} for class {ClassId} by user {UserId} | " +
                    "GradingType: {GradingType}, Endpoint: {Endpoint}",
                    assignment.Name, assignment.ClassId, userId,
                    assignment.GradingType, assignment.GradingApiEndpoint ?? "N/A");

                return assignment;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error creating assignment");
                throw;
            }
        }

    }
}
