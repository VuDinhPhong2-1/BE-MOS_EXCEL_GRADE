using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class ExamPublicationService : IExamPublicationService
    {
        private readonly IMongoCollection<ExamPublication> _examPublications;
        private readonly IMongoCollection<Student> _students;
        private readonly ILogger<ExamPublicationService> _logger;

        public ExamPublicationService(
            IMongoDatabase database,
            ILogger<ExamPublicationService> logger)
        {
            _examPublications = database.GetCollection<ExamPublication>("examPublications");
            _students = database.GetCollection<Student>("students");
            _logger = logger;
        }

        public async Task<ExamPublication> CreateExamPublicationAsync(CreateExamPublicationRequest request, string userId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.Name))
                {
                    throw new ArgumentException("Tên ca thi là bắt buộc.");
                }

                request.StudentIds ??= new List<string>();
                request.ProjectSequence ??= new List<CreateExamPublicationProjectRequest>();
                request.TaskSnapshot ??= new List<ExamPublicationTaskSnapshotItemDto>();

                EnsureValidOptionalObjectId(request.ClassId, "Mã lớp");
                EnsureValidObjectIds(request.StudentIds, "StudentIds");

                if (request.DurationMinutes.HasValue && request.DurationMinutes <= 0)
                {
                    throw new ArgumentException("DurationMinutes phải lớn hơn 0.");
                }

                if (request.StartsAt.HasValue && request.EndsAt.HasValue && request.EndsAt < request.StartsAt)
                {
                    throw new ArgumentException("EndsAt phải lớn hơn hoặc bằng StartsAt.");
                }

                var normalizedProjectSequence = NormalizeProjectSequence(request);
                if (normalizedProjectSequence.Count == 0)
                {
                    throw new ArgumentException("Cần ít nhất một project trong projectSequence.");
                }

                var publication = new ExamPublication
                {
                    Name = request.Name.Trim(),
                    Description = string.IsNullOrWhiteSpace(request.Description) ? null : request.Description.Trim(),
                    ClassId = string.IsNullOrWhiteSpace(request.ClassId) ? null : request.ClassId.Trim(),
                    StudentIds = request.StudentIds
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .Select(id => id.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList(),
                    Mode = NormalizeNullable(request.Mode),
                    StartsAt = NormalizeUtcDateTime(request.StartsAt),
                    EndsAt = NormalizeUtcDateTime(request.EndsAt),
                    DurationMinutes = request.DurationMinutes,
                    ProjectSequence = normalizedProjectSequence,
                    CreatedBy = userId,
                    CreatedAt = DateTime.UtcNow
                };

                await _examPublications.InsertOneAsync(publication);

                _logger.LogInformation(
                    "✅ Exam publication created: {PublicationId} ({ProjectCount} projects) by user {UserId}",
                    publication.Id,
                    publication.ProjectSequence.Count,
                    userId);

                return publication;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error creating exam publication");
                throw;
            }
        }

        public async Task<ExamPublication?> GetExamPublicationByIdAsync(string id)
        {
            try
            {
                EnsureRequiredObjectId(id, "Mã publication");
                return await _examPublications.Find(item => item.Id == id).FirstOrDefaultAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error getting exam publication {PublicationId}", id);
                throw;
            }
        }

        public async Task<PublicExamPublicationInfoDto?> GetPublicExamPublicationByTokenAsync(string publicationToken)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(publicationToken))
                {
                    throw new ArgumentException("Publication token là bắt buộc.");
                }

                var normalizedToken = publicationToken.Trim();
                var publication = await _examPublications
                    .Find(item => item.PublicationToken == normalizedToken && item.IsActive)
                    .FirstOrDefaultAsync();

                if (publication == null)
                {
                    return null;
                }

                var studentIds = publication.StudentIds
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .Select(id => id.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var students = new List<PublicExamStudentDto>();

                if (studentIds.Count > 0)
                {
                    var studentFilter = Builders<Student>.Filter.And(
                        Builders<Student>.Filter.In(student => student.Id, studentIds),
                        Builders<Student>.Filter.Eq(student => student.IsActive, true));

                    var matchedStudents = await _students
                        .Find(studentFilter)
                        .ToListAsync();

                    var studentLookup = matchedStudents
                        .Where(student => !string.IsNullOrWhiteSpace(student.Id))
                        .ToDictionary(
                        student => student.Id!,
                        student => new PublicExamStudentDto
                        {
                            Id = student.Id ?? string.Empty,
                            FullName = $"{student.MiddleName} {student.FirstName}".Trim()
                        },
                        StringComparer.OrdinalIgnoreCase);

                    students = studentIds
                        .Where(studentLookup.ContainsKey)
                        .Select(id => studentLookup[id])
                        .ToList();
                }

                return new PublicExamPublicationInfoDto
                {
                    Id = publication.Id,
                    Name = publication.Name,
                    Description = publication.Description,
                    StartsAt = publication.StartsAt,
                    EndsAt = publication.EndsAt,
                    DurationMinutes = publication.DurationMinutes,
                    StudentIds = studentIds,
                    Students = students,
                    ProjectCount = publication.ProjectSequence?.Count ?? 0
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting public exam publication by token");
                throw;
            }
        }

        private static List<ExamPublicationProject> NormalizeProjectSequence(CreateExamPublicationRequest request)
        {
            var requestedProjects = (request.ProjectSequence ?? new List<CreateExamPublicationProjectRequest>())
                .Where(item => item != null)
                .ToList();

            if (requestedProjects.Count == 0 && !string.IsNullOrWhiteSpace(request.GradingApiEndpoint))
            {
                requestedProjects.Add(new CreateExamPublicationProjectRequest
                {
                    Order = 1,
                    ProjectCode = request.ProjectCode,
                    Subject = request.Subject,
                    TemplateFileName = request.TemplateFileName,
                    GradingApiEndpoint = request.GradingApiEndpoint,
                    TaskSnapshot = request.TaskSnapshot,
                    ModeRules = request.ModeRules
                });
            }

            var normalizedProjects = new List<ExamPublicationProject>();

            for (var index = 0; index < requestedProjects.Count; index++)
            {
                var item = requestedProjects[index];

                if (string.IsNullOrWhiteSpace(item.GradingApiEndpoint))
                {
                    throw new ArgumentException($"ProjectSequence[{index}].GradingApiEndpoint là bắt buộc.");
                }

                var normalizedEndpoint = GradingApiEndpoints.NormalizeEndpoint(item.GradingApiEndpoint);
                if (!GradingApiEndpoints.IsValidEndpoint(normalizedEndpoint))
                {
                    throw new ArgumentException($"GradingApiEndpoint không hợp lệ: {item.GradingApiEndpoint}");
                }

                if (!TryInferSubject(normalizedEndpoint, out var subject))
                {
                    throw new ArgumentException($"Không xác định được subject từ endpoint: {item.GradingApiEndpoint}");
                }

                if (!GradingApiEndpoints.TryExtractProjectNumber(normalizedEndpoint, out var projectNumber))
                {
                    throw new ArgumentException($"Không xác định được project number từ endpoint: {item.GradingApiEndpoint}");
                }

                var normalizedSubject = AssignmentFileSubjects.Normalize(item.Subject ?? subject);
                if (!AssignmentFileSubjects.IsValid(normalizedSubject))
                {
                    throw new ArgumentException($"Subject không hợp lệ tại projectSequence[{index}].");
                }

                if (!string.Equals(normalizedSubject, subject, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException(
                        $"Subject không khớp với GradingApiEndpoint tại projectSequence[{index}].");
                }

                normalizedProjects.Add(new ExamPublicationProject
                {
                    Order = item.Order.GetValueOrDefault(index + 1),
                    ProjectCode = NormalizeProjectCode(item.ProjectCode, normalizedSubject, projectNumber),
                    Subject = normalizedSubject,
                    TemplateFileName = NormalizeNullable(item.TemplateFileName),
                    GradingApiEndpoint = normalizedEndpoint,
                    TaskSnapshot = (item.TaskSnapshot ?? new List<ExamPublicationTaskSnapshotItemDto>())
                        .Select(MapTaskSnapshot)
                        .ToList(),
                    ModeRules = MapModeRules(item.ModeRules)
                });
            }

            if (normalizedProjects.Any(item => item.Order <= 0))
            {
                throw new ArgumentException("Order trong projectSequence phải lớn hơn 0.");
            }

            var duplicateOrders = normalizedProjects
                .GroupBy(item => item.Order)
                .FirstOrDefault(group => group.Count() > 1);

            if (duplicateOrders != null)
            {
                throw new ArgumentException($"Order bị trùng trong projectSequence: {duplicateOrders.Key}");
            }

            return normalizedProjects
                .OrderBy(item => item.Order)
                .ToList();
        }

        private static bool TryInferSubject(string normalizedEndpoint, out string subject)
        {
            subject = string.Empty;
            var parts = normalizedEndpoint.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 2)
            {
                return false;
            }

            subject = AssignmentFileSubjects.Normalize(parts[0]);
            return AssignmentFileSubjects.IsValid(subject);
        }

        private static string NormalizeProjectCode(string? projectCode, string subject, int projectNumber)
        {
            if (!string.IsNullOrWhiteSpace(projectCode))
            {
                return projectCode.Trim().ToUpperInvariant();
            }

            return $"{subject.ToUpperInvariant()}_P{projectNumber:00}";
        }

        private static ExamPublicationTaskSnapshotItem MapTaskSnapshot(ExamPublicationTaskSnapshotItemDto item)
        {
            return new ExamPublicationTaskSnapshotItem
            {
                TaskId = NormalizeNullable(item.TaskId),
                TaskName = NormalizeNullable(item.TaskName),
                MaxScore = item.MaxScore,
                Instructions = NormalizeNullable(item.Instructions)
            };
        }

        private static ExamPublicationModeRules? MapModeRules(ExamPublicationModeRulesDto? item)
        {
            if (item == null)
            {
                return null;
            }

            return new ExamPublicationModeRules
            {
                Mode = NormalizeNullable(item.Mode),
                ShowFeedback = item.ShowFeedback,
                AllowRestart = item.AllowRestart,
                AllowNextProject = item.AllowNextProject
            };
        }

        private static string? NormalizeNullable(string? value) =>
            string.IsNullOrWhiteSpace(value) ? null : value.Trim();

        private static DateTime? NormalizeUtcDateTime(DateTime? value)
        {
            if (!value.HasValue)
            {
                return null;
            }

            return value.Value.Kind switch
            {
                DateTimeKind.Utc => value.Value,
                DateTimeKind.Local => value.Value.ToUniversalTime(),
                _ => DateTime.SpecifyKind(value.Value, DateTimeKind.Local).ToUniversalTime()
            };
        }

        private static void EnsureRequiredObjectId(string? value, string fieldLabel)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new ArgumentException($"{fieldLabel} không hợp lệ.");
            }
        }

        private static void EnsureValidOptionalObjectId(string? value, string fieldLabel)
        {
            if (!string.IsNullOrWhiteSpace(value) && !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new ArgumentException($"{fieldLabel} không hợp lệ.");
            }
        }

        private static void EnsureValidObjectIds(IEnumerable<string> values, string fieldLabel)
        {
            foreach (var value in values.Where(item => !string.IsNullOrWhiteSpace(item)))
            {
                if (!ObjectId.TryParse(value.Trim(), out _))
                {
                    throw new ArgumentException($"{fieldLabel} chứa ObjectId không hợp lệ: {value}");
                }
            }
        }
    }
}
