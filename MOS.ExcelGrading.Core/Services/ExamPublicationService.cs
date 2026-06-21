using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.GridFS;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using System.Text;

namespace MOS.ExcelGrading.Core.Services
{
    public class ExamPublicationService : IExamPublicationService
    {
        private static readonly Encoding Utf8Strict = new UTF8Encoding(
            encoderShouldEmitUTF8Identifier: false,
            throwOnInvalidBytes: true);

        private readonly IMongoCollection<ExamPublication> _examPublications;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<Assignment> _assignments;
        private readonly IMongoCollection<AssignmentFile> _assignmentFiles;
        private readonly GridFSBucket _bucket;
        private readonly IGradingService _gradingService;
        private readonly ILogger<ExamPublicationService> _logger;

        static ExamPublicationService()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public ExamPublicationService(
            IMongoDatabase database,
            IGradingService gradingService,
            ILogger<ExamPublicationService> logger)
        {
            _examPublications = database.GetCollection<ExamPublication>("examPublications");
            _students = database.GetCollection<Student>("students");
            _assignments = database.GetCollection<Assignment>("assignments");
            _assignmentFiles = database.GetCollection<AssignmentFile>("assignment_files");
            _bucket = new GridFSBucket(database, new GridFSBucketOptions
            {
                BucketName = "assignmentFiles"
            });
            _gradingService = gradingService;
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
                request.AssignmentIds ??= new List<string>();
                request.ProjectSequence ??= new List<CreateExamPublicationProjectRequest>();
                request.TaskSnapshot ??= new List<ExamPublicationTaskSnapshotItemDto>();

                EnsureValidOptionalObjectId(request.ClassId, "Mã lớp");
                EnsureValidObjectIds(request.StudentIds, "StudentIds");
                EnsureValidObjectIds(request.AssignmentIds, "AssignmentIds");

                if (request.DurationMinutes.HasValue && request.DurationMinutes <= 0)
                {
                    throw new ArgumentException("DurationMinutes phải lớn hơn 0.");
                }

                if (request.StartsAt.HasValue && request.EndsAt.HasValue && request.EndsAt < request.StartsAt)
                {
                    throw new ArgumentException("EndsAt phải lớn hơn hoặc bằng StartsAt.");
                }

                var normalizedProjectSequence = await NormalizeProjectSequenceAsync(request);
                if (normalizedProjectSequence.Count == 0)
                {
                    throw new ArgumentException("Cần ít nhất một project trong lịch thi.");
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
                    AssignmentIds = normalizedProjectSequence
                        .Select(item => item.SourceAssignmentId)
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList()!,
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

        private async Task<List<ExamPublicationProject>> NormalizeProjectSequenceAsync(CreateExamPublicationRequest request)
        {
            if (request.AssignmentIds.Count > 0)
            {
                return await NormalizeProjectSequenceFromAssignmentsAsync(request);
            }

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
                    SourceAssignmentId = NormalizeNullable(item.SourceAssignmentId),
                    ProjectCode = NormalizeProjectCode(item.ProjectCode, normalizedSubject, projectNumber),
                    Subject = normalizedSubject,
                    TemplateFileName = NormalizeNullable(item.TemplateFileName),
                    InstructionsFileName = NormalizeNullable(item.InstructionsFileName),
                    InstructionsText = NormalizeLongText(item.InstructionsText),
                    HelpFileName = NormalizeNullable(item.HelpFileName),
                    HelpText = NormalizeLongText(item.HelpText),
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

        private async Task<List<ExamPublicationProject>> NormalizeProjectSequenceFromAssignmentsAsync(CreateExamPublicationRequest request)
        {
            var assignmentIds = request.AssignmentIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id.Trim())
                .ToList();

            if (assignmentIds.Count != assignmentIds.Distinct(StringComparer.OrdinalIgnoreCase).Count())
            {
                throw new ArgumentException("Không được chọn trùng assignment trong cùng một lịch thi.");
            }

            var assignments = await _assignments
                .Find(item => assignmentIds.Contains(item.Id) && item.IsActive)
                .ToListAsync();

            if (assignments.Count != assignmentIds.Count)
            {
                throw new ArgumentException("Danh sách assignment chứa bài tập không tồn tại hoặc đã ngừng hoạt động.");
            }

            var assignmentById = assignments.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
            var orderedAssignments = assignmentIds.Select(id => assignmentById[id]).ToList();

            if (!string.IsNullOrWhiteSpace(request.ClassId) &&
                orderedAssignments.Any(item => !string.Equals(item.ClassId, request.ClassId, StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException("Tất cả assignment phải thuộc đúng lớp được chọn.");
            }

            var distinctClassIds = orderedAssignments
                .Select(item => item.ClassId)
                .Where(classId => !string.IsNullOrWhiteSpace(classId))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (distinctClassIds.Count > 1)
            {
                throw new ArgumentException("Các assignment được chọn phải cùng thuộc một lớp.");
            }

            if (string.IsNullOrWhiteSpace(request.ClassId) && distinctClassIds.Count == 1)
            {
                request.ClassId = distinctClassIds[0];
            }

            var templateIds = orderedAssignments
                .Where(item => !string.IsNullOrWhiteSpace(item.CurrentTemplateFileId))
                .Select(item => item.CurrentTemplateFileId!)
                .ToList();

            var instructionsIds = orderedAssignments
                .Where(item => !string.IsNullOrWhiteSpace(item.CurrentInstructionsFileId))
                .Select(item => item.CurrentInstructionsFileId!)
                .ToList();

            var helpIds = orderedAssignments
                .Where(item => !string.IsNullOrWhiteSpace(item.CurrentHelpFileId))
                .Select(item => item.CurrentHelpFileId!)
                .ToList();

            var fileIds = templateIds
                .Concat(instructionsIds)
                .Concat(helpIds)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var filesById = fileIds.Count == 0
                ? new Dictionary<string, AssignmentFile>(StringComparer.OrdinalIgnoreCase)
                : (await _assignmentFiles.Find(file => fileIds.Contains(file.Id)).ToListAsync())
                    .ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);

            var result = new List<ExamPublicationProject>();
            for (var index = 0; index < orderedAssignments.Count; index++)
            {
                var assignment = orderedAssignments[index];

                if (string.IsNullOrWhiteSpace(assignment.ExamType) ||
                    string.IsNullOrWhiteSpace(assignment.Subject) ||
                    string.IsNullOrWhiteSpace(assignment.ProjectCode))
                {
                    throw new ArgumentException(
                        $"Assignment '{assignment.Name}' là dữ liệu cũ và đang thiếu metadata chấm điểm. Vui lòng cập nhật exam type, subject và project code trước khi tạo lịch thi.");
                }

                var (isPublishable, publishBlockReason) = AssignmentService.EvaluatePublicationEligibility(assignment);
                if (!isPublishable)
                {
                    throw new ArgumentException(
                        $"Assignment '{assignment.Name}' không thể dùng để tạo lịch thi. {publishBlockReason}");
                }

                var route = ResolveAssignmentRoute(assignment);

                if (string.Equals(route.ExamType, AssignmentExamTypes.GMetrix, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException($"Assignment '{assignment.Name}' thuộc loại GMetrix và chưa được hỗ trợ trong lịch thi/runtime hiện tại.");
                }

                if (!route.IsRuntimeSupported)
                {
                    throw new ArgumentException($"Assignment '{assignment.Name}' chưa có grader runtime được hỗ trợ.");
                }

                if (assignment.GradingType != GradingTypes.Auto || string.IsNullOrWhiteSpace(route.GradingApiEndpoint))
                {
                    throw new ArgumentException($"Assignment '{assignment.Name}' chưa có grading endpoint tự động hợp lệ.");
                }

                filesById.TryGetValue(assignment.CurrentTemplateFileId ?? string.Empty, out var templateFile);
                filesById.TryGetValue(assignment.CurrentInstructionsFileId ?? string.Empty, out var instructionsFile);
                filesById.TryGetValue(assignment.CurrentHelpFileId ?? string.Empty, out var helpFile);

                result.Add(new ExamPublicationProject
                {
                    Order = index + 1,
                    SourceAssignmentId = assignment.Id,
                    ProjectCode = route.ProjectCode,
                    Subject = route.Subject,
                    TemplateFileName = templateFile?.OriginalName,
                    InstructionsFileName = instructionsFile?.OriginalName,
                    InstructionsText = await ReadOptionalTextFileAsync(instructionsFile),
                    HelpFileName = helpFile?.OriginalName,
                    HelpText = await ReadOptionalTextFileAsync(helpFile),
                    GradingApiEndpoint = route.GradingApiEndpoint ?? string.Empty,
                    TaskSnapshot = _gradingService
                        .GetTaskSnapshotForEndpoint(route.GradingApiEndpoint ?? string.Empty)
                        .Select(MapTaskSnapshot)
                        .ToList(),
                    ModeRules = BuildDefaultModeRules(request.Mode)
                });
            }

            return result;
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

        private async Task<string?> ReadOptionalTextFileAsync(AssignmentFile? file)
        {
            if (file == null || string.IsNullOrWhiteSpace(file.GridFsFileId))
            {
                return null;
            }

            if (!ObjectId.TryParse(file.GridFsFileId, out var gridFsObjectId))
            {
                return null;
            }

            await using var stream = await _bucket.OpenDownloadStreamAsync(gridFsObjectId);
            using var memoryStream = new MemoryStream();
            await stream.CopyToAsync(memoryStream);

            var content = DecodeTextFile(memoryStream.ToArray());
            return NormalizeLongText(content);
        }

        private static string DecodeTextFile(byte[] bytes)
        {
            if (bytes.Length == 0)
            {
                return string.Empty;
            }

            var encoding = DetectEncodingFromBom(bytes);
            if (encoding != null)
            {
                return RepairMojibakeIfNeeded(encoding.GetString(RemoveBom(bytes)));
            }

            try
            {
                return RepairMojibakeIfNeeded(Utf8Strict.GetString(bytes));
            }
            catch (DecoderFallbackException)
            {
                return RepairMojibakeIfNeeded(Encoding.GetEncoding(1258).GetString(bytes));
            }
        }

        private static string RepairMojibakeIfNeeded(string value)
        {
            if (!LooksLikeUtf8ReadAsWindows1252(value))
            {
                return value;
            }

            try
            {
                return Encoding.UTF8.GetString(Encoding.GetEncoding(1252).GetBytes(value));
            }
            catch
            {
                return value;
            }
        }

        private static bool LooksLikeUtf8ReadAsWindows1252(string value)
        {
            return value.Contains('Ã') ||
                value.Contains('Ä') ||
                value.Contains('Æ') ||
                value.Contains("áº", StringComparison.Ordinal) ||
                value.Contains("á»", StringComparison.Ordinal);
        }

        private static Encoding? DetectEncodingFromBom(byte[] bytes)
        {
            if (bytes.Length >= 3 &&
                bytes[0] == 0xEF &&
                bytes[1] == 0xBB &&
                bytes[2] == 0xBF)
            {
                return Encoding.UTF8;
            }

            if (bytes.Length >= 4 &&
                bytes[0] == 0xFF &&
                bytes[1] == 0xFE &&
                bytes[2] == 0x00 &&
                bytes[3] == 0x00)
            {
                return Encoding.UTF32;
            }

            if (bytes.Length >= 2 &&
                bytes[0] == 0xFF &&
                bytes[1] == 0xFE)
            {
                return Encoding.Unicode;
            }

            if (bytes.Length >= 2 &&
                bytes[0] == 0xFE &&
                bytes[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode;
            }

            return null;
        }

        private static byte[] RemoveBom(byte[] bytes)
        {
            if (bytes.Length >= 3 &&
                bytes[0] == 0xEF &&
                bytes[1] == 0xBB &&
                bytes[2] == 0xBF)
            {
                return bytes[3..];
            }

            if (bytes.Length >= 4 &&
                bytes[0] == 0xFF &&
                bytes[1] == 0xFE &&
                bytes[2] == 0x00 &&
                bytes[3] == 0x00)
            {
                return bytes[4..];
            }

            if (bytes.Length >= 2 &&
                ((bytes[0] == 0xFF && bytes[1] == 0xFE) ||
                 (bytes[0] == 0xFE && bytes[1] == 0xFF)))
            {
                return bytes[2..];
            }

            return bytes;
        }

        private static string? NormalizeLongText(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            return value.Replace("\r\n", "\n", StringComparison.Ordinal).Trim();
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
                AllowNextProject = item.AllowNextProject,
                AllowHelp = item.AllowHelp
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

        private static GraderRouteDescriptor ResolveAssignmentRoute(Assignment assignment)
        {
            if (!GraderRouteRegistry.TryResolve(
                    assignment.ExamType,
                    assignment.Subject,
                    assignment.ProjectCode,
                    assignment.GradingApiEndpoint,
                    out var route))
            {
                throw new ArgumentException($"Assignment '{assignment.Name}' có grading route không hợp lệ.");
            }

            return route;
        }

        private static ExamPublicationModeRules BuildDefaultModeRules(string? mode)
        {
            var normalizedMode = NormalizeNullable(mode) ?? "Testing";
            var isTraining = string.Equals(normalizedMode, "Training", StringComparison.OrdinalIgnoreCase);

            return new ExamPublicationModeRules
            {
                Mode = normalizedMode,
                ShowFeedback = isTraining,
                AllowRestart = true,
                AllowNextProject = true,
                AllowHelp = isTraining
            };
        }
    }
}
