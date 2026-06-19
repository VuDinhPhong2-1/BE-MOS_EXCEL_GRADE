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
        private readonly IMongoCollection<ExamPublication> _examPublications;
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
            _examPublications = database.GetCollection<ExamPublication>("examPublications");
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

                var assignments = await _assignments.Find(filter)
                    .SortByDescending(a => a.CreatedAt)
                    .ToListAsync();
                await ApplyPublicationStatesAsync(assignments);
                return assignments;
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

            var normalizedAssignments = assignments
                .OrderByDescending(a => a.CreatedAt)
                .ToList();
            await ApplyPublicationStatesAsync(normalizedAssignments);
            return normalizedAssignments;
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
                    Subject = AssignmentFileSubjects.Normalize(ReadString(document, "subject")) switch
                    {
                        var normalized when AssignmentFileSubjects.IsValid(normalized) => normalized,
                        _ => AssignmentFileSubjects.Excel
                    },
                    ExamType = AssignmentExamTypes.IsValid(ReadString(document, "examType"))
                        ? AssignmentExamTypes.Normalize(ReadString(document, "examType"))
                        : AssignmentExamTypes.OTTH,
                    ProjectCode = ReadString(document, "projectCode"),
                    CurrentTemplateFileId = ReadString(document, "currentTemplateFileId"),
                    CurrentAnswerFileId = ReadString(document, "currentAnswerFileId"),
                    CurrentInstructionsFileId = ReadString(document, "currentInstructionsFileId"),
                    CurrentHelpFileId = ReadString(document, "currentHelpFileId"),
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
                        Subject = assignment.Subject,
                        ExamType = assignment.ExamType,
                        ProjectCode = assignment.ProjectCode,
                        CreatedAt = assignment.CreatedAt,
                        IsActive = assignment.IsActive,
                        IsLockedForPublication = assignment.IsLockedForPublication,
                        IsPublishable = assignment.IsPublishable,
                        PublishBlockReason = assignment.PublishBlockReason,
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
                var assignment = await _assignments.Find(a => a.Id == id).FirstOrDefaultAsync();
                if (assignment != null)
                {
                    await ApplyPublicationStatesAsync(new List<Assignment> { assignment });
                }

                return assignment;
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

                var existingAssignment = await _assignments.Find(a => a.Id == id).FirstOrDefaultAsync();
                if (existingAssignment == null)
                {
                    return null;
                }

                var assignmentUsedInPublication = await IsAssignmentUsedInPublicationAsync(id);
                var nextGradingType = string.IsNullOrWhiteSpace(request.GradingType)
                    ? existingAssignment.GradingType
                    : request.GradingType.Trim().ToLowerInvariant();
                var normalizedEndpoint = string.IsNullOrWhiteSpace(request.GradingApiEndpoint)
                    ? existingAssignment.GradingApiEndpoint
                    : GradingApiEndpoints.NormalizeEndpoint(request.GradingApiEndpoint);
                var nextExamType = string.IsNullOrWhiteSpace(request.ExamType)
                    ? existingAssignment.ExamType
                    : AssignmentExamTypes.Normalize(request.ExamType);
                var nextSubject = string.IsNullOrWhiteSpace(request.Subject)
                    ? existingAssignment.Subject
                    : AssignmentFileSubjects.Normalize(request.Subject);
                var nextProjectCode = string.IsNullOrWhiteSpace(request.ProjectCode)
                    ? existingAssignment.ProjectCode
                    : request.ProjectCode.Trim().ToUpperInvariant();

                if (!AssignmentExamTypes.IsValid(nextExamType))
                {
                    throw new ArgumentException("ExamType không hợp lệ. Chỉ chấp nhận: OTTH, OnThi, GMetrix.");
                }

                if (!AssignmentFileSubjects.IsValid(nextSubject))
                {
                    throw new ArgumentException("Subject không hợp lệ. Chỉ chấp nhận: excel, word, ppt.");
                }

                if (assignmentUsedInPublication &&
                    (nextGradingType != existingAssignment.GradingType ||
                     !string.Equals(normalizedEndpoint, existingAssignment.GradingApiEndpoint, StringComparison.OrdinalIgnoreCase) ||
                     !string.Equals(nextExamType, existingAssignment.ExamType, StringComparison.OrdinalIgnoreCase) ||
                     !string.Equals(nextSubject, existingAssignment.Subject, StringComparison.OrdinalIgnoreCase) ||
                     !string.Equals(nextProjectCode, existingAssignment.ProjectCode, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new ArgumentException("Bài tập đã được dùng để tạo lịch thi nên không thể sửa các trường lõi (loại đề, môn, project, grading endpoint).");
                }

                var route = ResolveRouteOrThrow(nextExamType, nextSubject, nextProjectCode, nextGradingType == GradingTypes.Auto ? normalizedEndpoint : null);
                normalizedEndpoint = nextGradingType == GradingTypes.Auto ? route.GradingApiEndpoint : null;
                nextSubject = route.Subject;
                nextProjectCode = route.ProjectCode;

                var updateDefinitions = new List<UpdateDefinition<Assignment>>();
                var builder = Builders<Assignment>.Update;

                if (!string.IsNullOrEmpty(request.Name))
                    updateDefinitions.Add(builder.Set(a => a.Name, request.Name));

                if (request.Description != null)
                    updateDefinitions.Add(builder.Set(a => a.Description, request.Description));

                if (request.MaxScore.HasValue)
                    updateDefinitions.Add(builder.Set(a => a.MaxScore, request.MaxScore.Value));

                if (!string.IsNullOrWhiteSpace(request.GradingType))
                    updateDefinitions.Add(builder.Set(a => a.GradingType, nextGradingType));

                if (request.GradingApiEndpoint != null)
                    updateDefinitions.Add(builder.Set(a => a.GradingApiEndpoint, normalizedEndpoint));

                if (request.Subject != null)
                    updateDefinitions.Add(builder.Set(a => a.Subject, nextSubject));

                if (request.ExamType != null)
                    updateDefinitions.Add(builder.Set(a => a.ExamType, nextExamType));

                if (request.ProjectCode != null)
                    updateDefinitions.Add(builder.Set(a => a.ProjectCode, nextProjectCode));

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
                    await ApplyPublicationStatesAsync(new List<Assignment> { result });
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
                EnsureValidObjectId(request.ClassId, "Mã lớp");

                if (request.MaxScore < 0 || request.MaxScore > 1000)
                {
                    throw new ArgumentException("Điểm tối đa phải nằm trong khoảng từ 0 đến 1000");
                }

                var normalizedGradingType = (request.GradingType ?? string.Empty).Trim().ToLowerInvariant();
                if (normalizedGradingType != GradingTypes.Auto &&
                    normalizedGradingType != GradingTypes.Manual)
                {
                    throw new ArgumentException("GradingType phải là 'auto' hoặc 'manual'");
                }

                if (!AssignmentExamTypes.IsValid(request.ExamType))
                {
                    throw new ArgumentException("ExamType không hợp lệ. Chỉ chấp nhận: OTTH, OnThi, GMetrix.");
                }

                var normalizedSubject = AssignmentFileSubjects.Normalize(request.Subject);
                if (!AssignmentFileSubjects.IsValid(normalizedSubject))
                {
                    throw new ArgumentException("Subject không hợp lệ. Chỉ chấp nhận: excel, word, ppt.");
                }

                // ✅ VALIDATE GRADING API ENDPOINT
                if (normalizedGradingType == GradingTypes.Auto)
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

                var route = ResolveRouteOrThrow(
                    request.ExamType,
                    normalizedSubject,
                    request.ProjectCode,
                    normalizedGradingType == GradingTypes.Auto ? request.GradingApiEndpoint : null);

                var assignment = new Assignment
                {
                    Name = request.Name,
                    Description = request.Description,
                    ClassId = request.ClassId,
                    MaxScore = request.MaxScore,
                    GradingType = normalizedGradingType,
                    GradingApiEndpoint = route.GradingApiEndpoint,
                    Subject = route.Subject,
                    ExamType = route.ExamType,
                    ProjectCode = route.ProjectCode,
                    CreatedBy = userId,
                    CreatedAt = DateTime.UtcNow
                };

                await _assignments.InsertOneAsync(assignment);
                await ApplyPublicationStatesAsync(new List<Assignment> { assignment });

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

        public Task<List<AssignmentTemplateResponse>> GetAssignmentTemplatesAsync(string classId, string subject, string examType)
        {
            EnsureValidObjectId(classId, "Mã lớp");

            var normalizedSubject = AssignmentFileSubjects.Normalize(subject);
            var normalizedExamType = AssignmentExamTypes.Normalize(examType);

            if (!AssignmentFileSubjects.IsValid(normalizedSubject))
            {
                throw new ArgumentException("Subject không hợp lệ. Chỉ chấp nhận: excel, word, ppt.");
            }

            if (!AssignmentExamTypes.IsValid(normalizedExamType))
            {
                throw new ArgumentException("ExamType không hợp lệ. Chỉ chấp nhận: OTTH, OnThi, GMetrix.");
            }

            if (normalizedExamType != AssignmentExamTypes.OnThi)
            {
                return Task.FromResult(new List<AssignmentTemplateResponse>());
            }

            var projectNumbers = normalizedSubject switch
            {
                AssignmentFileSubjects.Excel => AssignmentTemplateRules.OnThiExcelProjectNumbers,
                AssignmentFileSubjects.Word => AssignmentTemplateRules.OnThiWordProjectNumbers,
                _ => Array.Empty<int>()
            };

            return Task.FromResult(projectNumbers
                .Select(projectNumber =>
                {
                    var endpoint = GradingApiEndpoints.ToProjectEndpoint(normalizedSubject, projectNumber);
                    var practice = PracticeScoring.ResolveByProjectNumber(projectNumber);
                    var route = ResolveRouteOrThrow(normalizedExamType, normalizedSubject, null, endpoint);

                    return new AssignmentTemplateResponse
                    {
                        SuggestedName = $"Ôn thi - Project {projectNumber:00} - {GetSubjectDisplayName(normalizedSubject)}",
                        Description = $"Bài ôn thi dùng grader OTTH cho Project {projectNumber:00}.",
                        Subject = route.Subject,
                        ExamType = normalizedExamType,
                        ProjectCode = route.ProjectCode,
                        GradingType = GradingTypes.Auto,
                        GradingApiEndpoint = route.GradingApiEndpoint,
                        MaxScore = (double)PracticeScoring.CalculateProjectMaxScore(projectNumber),
                        PracticeCode = practice.Code,
                        PracticeName = practice.Name
                    };
                })
                .ToList());
        }

        private async Task ApplyPublicationLocksAsync(List<Assignment> assignments)
        {
            if (assignments.Count == 0)
            {
                return;
            }

            var assignmentIds = assignments
                .Where(item => !string.IsNullOrWhiteSpace(item.Id))
                .Select(item => item.Id)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var publications = await _examPublications.Find(_ => true).ToListAsync();
            var lockedIds = publications
                .SelectMany(publication => publication.ProjectSequence ?? new List<ExamPublicationProject>())
                .Select(item => item.SourceAssignmentId)
                .Where(id => !string.IsNullOrWhiteSpace(id) && assignmentIds.Contains(id))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var assignment in assignments)
            {
                assignment.IsLockedForPublication = !string.IsNullOrWhiteSpace(assignment.Id) &&
                                                   lockedIds.Contains(assignment.Id);
                ApplyPublicationEligibility(assignment);
            }
        }

        private async Task ApplyPublicationStatesAsync(List<Assignment> assignments)
        {
            await ApplyPublicationLocksAsync(assignments);
        }

        private async Task<bool> IsAssignmentUsedInPublicationAsync(string assignmentId)
        {
            var filter = Builders<ExamPublication>.Filter.ElemMatch(
                publication => publication.ProjectSequence,
                project => project.SourceAssignmentId == assignmentId);

            return await _examPublications.Find(filter).AnyAsync();
        }

        private static GraderRouteDescriptor ResolveRouteOrThrow(
            string examType,
            string subject,
            string? projectCode,
            string? gradingApiEndpoint)
        {
            if (!GraderRouteRegistry.TryResolve(examType, subject, projectCode, gradingApiEndpoint, out var route))
            {
                throw new ArgumentException("Không thể xác định grading route hợp lệ cho assignment.");
            }

            return route;
        }

        private static void ApplyPublicationEligibility(Assignment assignment)
        {
            var (isPublishable, reason) = EvaluatePublicationEligibility(assignment);
            assignment.IsPublishable = isPublishable;
            assignment.PublishBlockReason = reason;
        }

        public static (bool IsPublishable, string? Reason) EvaluatePublicationEligibility(Assignment assignment)
        {
            if (!assignment.IsActive)
            {
                return (false, "Bài tập đã ngừng hoạt động nên không thể dùng để tạo lịch thi.");
            }

            if (assignment.GradingType != GradingTypes.Auto)
            {
                return (false, "Bài tập chấm thủ công chưa thể dùng để tạo lịch thi tự động.");
            }

            if (string.IsNullOrWhiteSpace(assignment.ExamType) ||
                string.IsNullOrWhiteSpace(assignment.Subject) ||
                string.IsNullOrWhiteSpace(assignment.ProjectCode))
            {
                return (false, "Bài tập cũ đang thiếu metadata chấm điểm (examType, subject hoặc projectCode).");
            }

            if (string.IsNullOrWhiteSpace(assignment.GradingApiEndpoint))
            {
                return (false, "Bài tập chưa có grading endpoint tự động hợp lệ.");
            }

            if (!AssignmentExamTypes.IsValid(assignment.ExamType))
            {
                return (false, "Loại đề của bài tập không hợp lệ.");
            }

            if (!AssignmentFileSubjects.IsValid(assignment.Subject))
            {
                return (false, "Môn của bài tập không hợp lệ.");
            }

            if (!GraderRouteRegistry.TryResolve(
                    assignment.ExamType,
                    assignment.Subject,
                    assignment.ProjectCode,
                    assignment.GradingApiEndpoint,
                    out var route))
            {
                return (false, "Metadata grading của bài tập không hợp lệ hoặc không khớp nhau.");
            }

            if (string.Equals(route.ExamType, AssignmentExamTypes.GMetrix, StringComparison.OrdinalIgnoreCase))
            {
                return (false, "GMetrix chưa được hỗ trợ trong runtime tạo lịch thi hiện tại.");
            }

            if (!route.IsRuntimeSupported)
            {
                return (false, "Loại đề của bài tập chưa được runtime hiện tại hỗ trợ.");
            }

            return (true, null);
        }

        private static string GetSubjectDisplayName(string subject) =>
            AssignmentFileSubjects.Normalize(subject) switch
            {
                AssignmentFileSubjects.Word => "Word",
                AssignmentFileSubjects.Ppt => "PowerPoint",
                _ => "Excel"
            };

    }
}

