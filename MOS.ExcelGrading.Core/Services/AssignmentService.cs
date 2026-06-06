// MOS.ExcelGrading.Core/Services/AssignmentService.cs
using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class AssignmentService : IAssignmentService
    {
        private readonly IMongoCollection<Assignment> _assignments;
        private readonly IMongoCollection<BsonDocument> _assignmentDocuments;
        private readonly IMongoCollection<Score> _scores;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<User> _users;
        private readonly ILogger<AssignmentService> _logger;

        public AssignmentService(
            IMongoDatabase database,
            ILogger<AssignmentService> logger)
        {
            _assignments = database.GetCollection<Assignment>("assignments");
            _assignmentDocuments = database.GetCollection<BsonDocument>("assignments");
            _scores = database.GetCollection<Score>("scores");
            _students = database.GetCollection<Student>("students");
            _users = database.GetCollection<User>("users");
            _logger = logger;
        }

        public async Task<List<Assignment>> GetAssignmentsByClassIdAsync(string classId, bool includeInactive = false)
        {
            try
            {
                EnsureValidObjectId(classId, "Mã lớp");

                var filter = Builders<Assignment>.Filter.Eq(a => a.ClassId, classId);
                if (!includeInactive)
                {
                    filter &= Builders<Assignment>.Filter.Eq(a => a.IsActive, true);
                }

                return await _assignments.Find(filter)
                    .SortByDescending(a => a.CreatedAt)
                    .ToListAsync();
            }
            catch (BsonSerializationException ex)
            {
                _logger.LogWarning(
                    ex,
                    "⚠️ Deserialization error while reading assignments for class {ClassId}. Falling back to tolerant document mapping.",
                    classId);

                return await GetAssignmentsByClassIdFromRawDocumentsAsync(classId, includeInactive);
            }
            catch (FormatException ex)
            {
                _logger.LogWarning(
                    ex,
                    "⚠️ Invalid assignment document format while reading assignments for class {ClassId}. Falling back to tolerant document mapping.",
                    classId);

                return await GetAssignmentsByClassIdFromRawDocumentsAsync(classId, includeInactive);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting assignments for class {ClassId}", classId);
                throw;
            }
        }

        private async Task<List<Assignment>> GetAssignmentsByClassIdFromRawDocumentsAsync(string classId, bool includeInactive)
        {
            var classIdFilter = Builders<BsonDocument>.Filter.Eq("classId", classId);
            if (ObjectId.TryParse(classId, out var classObjectId))
            {
                classIdFilter |= Builders<BsonDocument>.Filter.Eq("classId", classObjectId);
            }

            var filter = classIdFilter;
            if (!includeInactive)
            {
                filter &= Builders<BsonDocument>.Filter.Eq("isActive", true);
            }

            var documents = await _assignmentDocuments.Find(filter).ToListAsync();
            var assignments = new List<Assignment>();

            foreach (var document in documents)
            {
                if (TryMapAssignmentDocument(document, out var assignment))
                {
                    assignments.Add(assignment);
                    continue;
                }

                var documentId = ReadString(document, "_id") ?? "(unknown)";
                _logger.LogWarning(
                    "⚠️ Skipping malformed assignment document {AssignmentId} for class {ClassId}",
                    documentId,
                    classId);
            }

            return assignments
                .OrderByDescending(a => a.CreatedAt)
                .ToList();
        }

        private static bool TryMapAssignmentDocument(BsonDocument document, out Assignment assignment)
        {
            assignment = new Assignment();

            try
            {
                var id = ReadString(document, "_id");
                var classId = ReadString(document, "classId");
                var name = ReadString(document, "name");

                if (string.IsNullOrWhiteSpace(id) ||
                    string.IsNullOrWhiteSpace(classId) ||
                    string.IsNullOrWhiteSpace(name))
                {
                    return false;
                }

                assignment = new Assignment
                {
                    Id = id,
                    Name = name,
                    Description = ReadString(document, "description"),
                    ClassId = classId,
                    MaxScore = ReadDouble(document, "maxScore") ?? 10d,
                    GradingApiEndpoint = ReadString(document, "gradingApiEndpoint"),
                    GradingType = ReadString(document, "gradingType") ?? GradingTypes.Manual,
                    CurrentTemplateFileId = ReadString(document, "currentTemplateFileId"),
                    CurrentAnswerFileId = ReadString(document, "currentAnswerFileId"),
                    IsActive = ReadBool(document, "isActive") ?? true,
                    CreatedAt = ReadDateTime(document, "createdAt") ?? DateTime.UtcNow,
                    CreatedBy = ReadString(document, "createdBy"),
                    UpdatedAt = ReadDateTime(document, "updatedAt"),
                    UpdatedBy = ReadString(document, "updatedBy")
                };

                return true;
            }
            catch
            {
                assignment = new Assignment();
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

        private static bool? ReadBool(BsonDocument document, string fieldName)
        {
            if (!document.TryGetValue(fieldName, out var value) || value.IsBsonNull)
            {
                return null;
            }

            return value.BsonType switch
            {
                BsonType.Boolean => value.AsBoolean,
                BsonType.Int32 => value.AsInt32 != 0,
                BsonType.Int64 => value.AsInt64 != 0,
                BsonType.String when bool.TryParse(value.AsString, out var parsed) => parsed,
                _ => null
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

        private static void EnsureValidObjectId(string? value, string fieldLabel)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new ArgumentException($"{fieldLabel} không hợp lệ.");
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
                if (request.MaxScore.HasValue &&
                    (request.MaxScore.Value < 0 || request.MaxScore.Value > 1000))
                {
                    throw new ArgumentException("Điểm tối đa phải nằm trong khoảng từ 0 đến 1000");
                }

                if (!string.IsNullOrWhiteSpace(request.GradingType) &&
                    request.GradingType != GradingTypes.Auto &&
                    request.GradingType != GradingTypes.Manual)
                {
                    throw new ArgumentException("GradingType phải là 'auto' hoặc 'manual'");
                }

                if (!string.IsNullOrWhiteSpace(request.GradingType) &&
                    request.GradingType == GradingTypes.Auto &&
                    string.IsNullOrWhiteSpace(request.GradingApiEndpoint))
                {
                    throw new ArgumentException("Bắt buộc có GradingApiEndpoint khi GradingType là 'auto'");
                }

                if (!string.IsNullOrWhiteSpace(request.GradingApiEndpoint) &&
                    !GradingApiEndpoints.IsValidEndpoint(request.GradingApiEndpoint))
                {
                    throw new ArgumentException($"GradingApiEndpoint không hợp lệ: {request.GradingApiEndpoint}");
                }

                var normalizedEndpoint = string.IsNullOrWhiteSpace(request.GradingApiEndpoint)
                    ? null
                    : GradingApiEndpoints.NormalizeEndpoint(request.GradingApiEndpoint);

                var updateDefinitions = new List<UpdateDefinition<Assignment>>();
                var builder = Builders<Assignment>.Update;

                if (!string.IsNullOrEmpty(request.Name))
                    updateDefinitions.Add(builder.Set(a => a.Name, request.Name));

                if (request.Description != null)
                    updateDefinitions.Add(builder.Set(a => a.Description, request.Description));

                if (request.MaxScore.HasValue)
                    updateDefinitions.Add(builder.Set(a => a.MaxScore, request.MaxScore.Value));

                if (!string.IsNullOrWhiteSpace(request.GradingType))
                    updateDefinitions.Add(builder.Set(a => a.GradingType, request.GradingType));

                if (request.GradingApiEndpoint != null)
                    updateDefinitions.Add(builder.Set(a => a.GradingApiEndpoint, normalizedEndpoint));

                // If switching to manual grading, clear endpoint.
                if (request.GradingType == GradingTypes.Manual)
                    updateDefinitions.Add(builder.Set(a => a.GradingApiEndpoint, null));

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
                if (request.MaxScore < 0 || request.MaxScore > 1000)
                {
                    throw new ArgumentException("Điểm tối đa phải nằm trong khoảng từ 0 đến 1000");
                }

                if (request.GradingType != GradingTypes.Auto &&
                    request.GradingType != GradingTypes.Manual)
                {
                    throw new ArgumentException("GradingType phải là 'auto' hoặc 'manual'");
                }

                // ✅ VALIDATE GRADING API ENDPOINT
                if (request.GradingType == GradingTypes.Auto)
                {
                    if (string.IsNullOrEmpty(request.GradingApiEndpoint))
                    {
                        throw new ArgumentException("Bắt buộc có GradingApiEndpoint khi GradingType là 'auto'");
                    }

                    if (!GradingApiEndpoints.IsValidEndpoint(request.GradingApiEndpoint))
                    {
                        throw new ArgumentException($"GradingApiEndpoint không hợp lệ: {request.GradingApiEndpoint}");
                    }

                    request.GradingApiEndpoint = GradingApiEndpoints.NormalizeEndpoint(request.GradingApiEndpoint);
                }
                else
                {
                    request.GradingApiEndpoint = null;
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

