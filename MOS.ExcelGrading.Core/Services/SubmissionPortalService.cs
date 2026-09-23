using System.Security.Cryptography;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class SubmissionPortalService : ISubmissionPortalService
    {
        private readonly IMongoCollection<SubmissionPortal> _portals;
        private readonly IMongoCollection<SubmissionLog> _logs;
        private readonly IMongoCollection<SubmissionAlert> _alerts;
        private readonly IMongoCollection<Class> _classes;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<Assignment> _assignments;
        private readonly IMongoCollection<Score> _scores;
        private readonly IXmlGradingRuleService _xmlGradingRuleService;
        private readonly IScoreService _scoreService;
        private readonly ILogger<SubmissionPortalService> _logger;

        public SubmissionPortalService(
            IMongoDatabase database,
            IXmlGradingRuleService xmlGradingRuleService,
            IScoreService scoreService,
            ILogger<SubmissionPortalService> logger)
        {
            _portals = database.GetCollection<SubmissionPortal>("submission_portals");
            _logs = database.GetCollection<SubmissionLog>("submission_logs");
            _alerts = database.GetCollection<SubmissionAlert>("submission_alerts");
            _classes = database.GetCollection<Class>("Classes");
            _students = database.GetCollection<Student>("students");
            _assignments = database.GetCollection<Assignment>("assignments");
            _scores = database.GetCollection<Score>("scores");
            _xmlGradingRuleService = xmlGradingRuleService;
            _scoreService = scoreService;
            _logger = logger;
        }

        public async Task<SubmissionPortalResponse> CreateAsync(CreateSubmissionPortalRequest request, string userId)
        {
            ValidateRequest(request);
            var portal = new SubmissionPortal
            {
                Title = request.Title.Trim(),
                Description = request.Description?.Trim(),
                ClassIds = request.ClassIds.Distinct().ToList(),
                AssignmentIds = request.AssignmentIds.Distinct().ToList(),
                StartsAt = ToUtc(request.StartsAt),
                EndsAt = ToUtc(request.EndsAt),
                MaxSubmissionsPerStudent = Math.Max(0, request.MaxSubmissionsPerStudent),
                ScoringPolicy = SubmissionPortalScoringPolicies.Normalize(request.ScoringPolicy),
                ShowLeaderboard = request.ShowLeaderboard,
                ShowDetailedFeedback = request.ShowDetailedFeedback,
                CreatedBy = userId,
                CreatedAt = DateTime.UtcNow,
                IsActive = true
            };

            await EnsureReferencesAsync(portal.ClassIds, portal.AssignmentIds);
            await _portals.InsertOneAsync(portal);
            return await MapPortalAsync(portal);
        }

        public async Task<List<SubmissionPortalResponse>> GetAllAsync(string userId, bool isAdmin)
        {
            var filter = isAdmin ? Builders<SubmissionPortal>.Filter.Empty : Builders<SubmissionPortal>.Filter.Eq(p => p.CreatedBy, userId);
            var portals = await _portals.Find(filter).SortByDescending(p => p.CreatedAt).ToListAsync();
            var result = new List<SubmissionPortalResponse>();
            foreach (var portal in portals) result.Add(await MapPortalAsync(portal));
            return result;
        }

        public async Task<SubmissionPortalResponse?> GetByIdAsync(string id)
        {
            EnsureObjectId(id, "Mã link nộp bài");
            var portal = await _portals.Find(p => p.Id == id).FirstOrDefaultAsync();
            return portal == null ? null : await MapPortalAsync(portal);
        }

        public async Task<SubmissionPortalResponse?> UpdateAsync(string id, UpdateSubmissionPortalRequest request, string userId)
        {
            EnsureObjectId(id, "Mã link nộp bài");
            ValidateRequest(request);
            await EnsureReferencesAsync(request.ClassIds.Distinct().ToList(), request.AssignmentIds.Distinct().ToList());

            var update = Builders<SubmissionPortal>.Update
                .Set(p => p.Title, request.Title.Trim())
                .Set(p => p.Description, request.Description?.Trim())
                .Set(p => p.ClassIds, request.ClassIds.Distinct().ToList())
                .Set(p => p.AssignmentIds, request.AssignmentIds.Distinct().ToList())
                .Set(p => p.StartsAt, ToUtc(request.StartsAt))
                .Set(p => p.EndsAt, ToUtc(request.EndsAt))
                .Set(p => p.MaxSubmissionsPerStudent, Math.Max(0, request.MaxSubmissionsPerStudent))
                .Set(p => p.ScoringPolicy, SubmissionPortalScoringPolicies.Normalize(request.ScoringPolicy))
                .Set(p => p.ShowLeaderboard, request.ShowLeaderboard)
                .Set(p => p.ShowDetailedFeedback, request.ShowDetailedFeedback)
                .Set(p => p.IsActive, request.IsActive)
                .Set(p => p.UpdatedAt, DateTime.UtcNow);

            var portal = await _portals.FindOneAndUpdateAsync(p => p.Id == id, update, new FindOneAndUpdateOptions<SubmissionPortal> { ReturnDocument = ReturnDocument.After });
            return portal == null ? null : await MapPortalAsync(portal);
        }

        public async Task<bool> DeleteAsync(string id, string userId)
        {
            EnsureObjectId(id, "Mã link nộp bài");
            var result = await _portals.UpdateOneAsync(p => p.Id == id, Builders<SubmissionPortal>.Update.Set(p => p.IsActive, false).Set(p => p.UpdatedAt, DateTime.UtcNow));
            return result.ModifiedCount > 0;
        }

        public async Task<PublicPortalInfoResponse?> GetPublicInfoAsync(string token)
        {
            var portal = await FindActiveByTokenAsync(token);
            if (portal == null) return null;
            var classes = await _classes.Find(c => c.Id != null && portal.ClassIds.Contains(c.Id) && c.IsActive).ToListAsync();
            var assignments = await _assignments.Find(a => portal.AssignmentIds.Contains(a.Id) && a.IsActive).ToListAsync();
            return new PublicPortalInfoResponse
            {
                Id = portal.Id,
                Title = portal.Title,
                Description = portal.Description,
                StartsAt = portal.StartsAt,
                EndsAt = portal.EndsAt,
                MaxSubmissionsPerStudent = portal.MaxSubmissionsPerStudent,
                ScoringPolicy = portal.ScoringPolicy,
                ShowLeaderboard = portal.ShowLeaderboard,
                ShowDetailedFeedback = portal.ShowDetailedFeedback,
                Classes = classes.Select(c => new PublicPortalClassDto { Id = c.Id ?? string.Empty, Name = c.Name }).ToList(),
                Assignments = assignments.Select(MapPublicAssignment).ToList()
            };
        }

        public async Task<List<PublicPortalStudentDto>> GetPublicStudentsAsync(string token, string classId)
        {
            var portal = await RequireOpenPortalAsync(token, validateTime: false);
            if (!portal.ClassIds.Contains(classId)) throw new InvalidOperationException("Lớp không thuộc link nộp bài này.");
            var students = await _students.Find(s => s.ClassId == classId && s.IsActive).SortBy(s => s.MiddleName).ThenBy(s => s.FirstName).ToListAsync();
            return students.Select(s => new PublicPortalStudentDto { Id = s.Id ?? string.Empty, ClassId = classId, FullName = FullName(s) }).ToList();
        }

        public async Task<PublicPortalSubmitResult> GradeAndSubmitAsync(string token, string classId, string studentId, string assignmentId, IFormFile file, string? ipAddress, string? userAgent)
        {
            var portal = await RequireOpenPortalAsync(token, validateTime: true);
            if (file == null || file.Length == 0) throw new InvalidOperationException("Vui lòng chọn file bài làm.");
            await ValidateSubmissionScopeAsync(portal, classId, studentId, assignmentId);

            var assignment = await _assignments.Find(a => a.Id == assignmentId && a.IsActive).FirstOrDefaultAsync()
                ?? throw new InvalidOperationException("Không tìm thấy bài tập.");
            if (string.IsNullOrWhiteSpace(assignment.GradingApiEndpoint)) throw new InvalidOperationException("Bài tập chưa cấu hình API chấm tự động.");

            var previousLogs = await _logs.Find(l => l.PortalId == portal.Id && l.StudentId == studentId && l.AssignmentId == assignmentId).ToListAsync();
            if (portal.MaxSubmissionsPerStudent > 0 && previousLogs.Count >= portal.MaxSubmissionsPerStudent)
            {
                throw new InvalidOperationException($"Bạn đã nộp đủ {portal.MaxSubmissionsPerStudent} lần cho bài này.");
            }

            await using var memory = new MemoryStream();
            await file.CopyToAsync(memory);
            var hash = Convert.ToHexString(SHA256.HashData(memory.ToArray())).ToLowerInvariant();
            memory.Position = 0;

            var (subject, projectCode) = ResolveEndpoint(assignment.GradingApiEndpoint);
            var grading = await _xmlGradingRuleService.GradeAsync(memory, subject, projectCode);
            var taskRequests = grading.TaskResults.Select(t => new AutoGradingTaskResultRequest
            {
                TaskId = t.TaskId,
                TaskName = t.TaskName,
                Score = (double)t.Score,
                MaxScore = (double)t.MaxScore,
                IsPassed = t.IsPassed,
                Details = t.Details,
                Errors = portal.ShowDetailedFeedback ? t.Errors : new List<string>(),
                FixActions = portal.ShowDetailedFeedback ? t.FixActions.ToList() : new List<string>(),
                DisplayIssues = portal.ShowDetailedFeedback ? t.DisplayIssues.Select(d => new AutoGradingDisplayIssueRequest { Heading = d.Heading, Message = d.Message, FixAction = d.FixAction }).ToList() : new List<AutoGradingDisplayIssueRequest>()
            }).ToList();

            var scoreValue = (double)grading.TotalScore;
            var shouldPersist = await ShouldPersistScoreAsync(portal, studentId, assignmentId, scoreValue);
            if (shouldPersist)
            {
                await _scoreService.CreateOrUpdateScoreAsync(new CreateScoreRequest
                {
                    StudentId = studentId,
                    AssignmentId = assignmentId,
                    ClassId = classId,
                    ScoreValue = scoreValue,
                    Feedback = $"Nộp qua public link: {portal.Title}",
                    AutoGradingErrors = portal.ShowDetailedFeedback ? grading.TaskResults.SelectMany(t => t.Errors).ToList() : new List<string>(),
                    AutoGradingTaskResults = taskRequests
                }, portal.CreatedBy ?? ObjectId.Empty.ToString());
            }

            var log = new SubmissionLog
            {
                PortalId = portal.Id,
                StudentId = studentId,
                ClassId = classId,
                AssignmentId = assignmentId,
                ScoreValue = scoreValue,
                MaxScore = (double)grading.MaxScore,
                IpAddress = ipAddress,
                UserAgent = userAgent,
                FileHash = hash,
                FileName = file.FileName,
                FileSizeBytes = file.Length,
                SubmittedAt = DateTime.UtcNow
            };
            await _logs.InsertOneAsync(log);
            log.Alerts = await DetectAlertsAsync(portal, log, previousLogs);
            if (log.Alerts.Count > 0)
            {
                await _logs.UpdateOneAsync(l => l.Id == log.Id, Builders<SubmissionLog>.Update.Set(l => l.Alerts, log.Alerts));
            }

            var leaderboard = await GetLeaderboardAsync(token, classId, assignmentId);
            var rank = leaderboard.FirstOrDefault(x => x.StudentId == studentId)?.Rank;
            return new PublicPortalSubmitResult
            {
                ScoreValue = scoreValue,
                MaxScore = (double)grading.MaxScore,
                Feedback = "Đã chấm và nộp bài thành công.",
                AutoGradingErrors = portal.ShowDetailedFeedback ? grading.TaskResults.SelectMany(t => t.Errors).ToList() : new List<string>(),
                AutoGradingTaskResults = taskRequests,
                SubmittedAt = log.SubmittedAt,
                Rank = rank,
                Alerts = log.Alerts
            };
        }

        public async Task<List<SubmissionLeaderboardItem>> GetLeaderboardAsync(string token, string? classId = null, string? assignmentId = null)
        {
            var portal = await RequireOpenPortalAsync(token, validateTime: false);
            if (!portal.ShowLeaderboard) throw new InvalidOperationException("Bảng xếp hạng chưa được bật cho link này.");
            if (!string.IsNullOrWhiteSpace(classId) && !portal.ClassIds.Contains(classId)) throw new InvalidOperationException("Lớp không thuộc link nộp bài này.");
            if (!string.IsNullOrWhiteSpace(assignmentId) && !portal.AssignmentIds.Contains(assignmentId)) throw new InvalidOperationException("Bài tập không thuộc link nộp bài này.");

            var filter = Builders<Score>.Filter.In(s => s.AssignmentId, portal.AssignmentIds);
            if (!string.IsNullOrWhiteSpace(classId)) filter &= Builders<Score>.Filter.Eq(s => s.ClassId, classId);
            if (!string.IsNullOrWhiteSpace(assignmentId)) filter &= Builders<Score>.Filter.Eq(s => s.AssignmentId, assignmentId);
            var scores = await _scores.Find(filter).ToListAsync();
            var scoreStudentIds = scores.Select(x => x.StudentId).Distinct().ToList();
            var studentFilter = Builders<Student>.Filter.Ne(s => s.Id, null) & Builders<Student>.Filter.In(s => s.Id, scoreStudentIds);
            if (!string.IsNullOrWhiteSpace(classId))
            {
                studentFilter |= Builders<Student>.Filter.Eq(s => s.ClassId, classId) & Builders<Student>.Filter.Eq(s => s.IsActive, true);
            }
            var students = await _students.Find(studentFilter).SortBy(s => s.MiddleName).ThenBy(s => s.FirstName).ToListAsync();
            var classes = await _classes.Find(c => c.Id != null && portal.ClassIds.Contains(c.Id)).ToListAsync();
            var assignments = await _assignments.Find(a => portal.AssignmentIds.Contains(a.Id)).ToListAsync();
            var logs = await _logs.Find(l => l.PortalId == portal.Id).ToListAsync();

            var rows = scores.Select(score =>
            {
                var student = students.FirstOrDefault(s => s.Id == score.StudentId);
                var cls = classes.FirstOrDefault(c => c.Id == score.ClassId);
                var assignment = assignments.FirstOrDefault(a => a.Id == score.AssignmentId);
                return new SubmissionLeaderboardItem
                {
                    StudentId = score.StudentId,
                    StudentName = student == null ? "(Không rõ)" : FullName(student),
                    ClassId = score.ClassId,
                    ClassName = cls?.Name ?? "(Không rõ)",
                    AssignmentId = string.IsNullOrWhiteSpace(assignmentId) ? null : score.AssignmentId,
                    AssignmentName = string.IsNullOrWhiteSpace(assignmentId) ? null : assignment?.Name,
                    ScoreValue = score.ScoreValue ?? 0,
                    MaxScore = assignment?.MaxScore ?? 125,
                    GradedAt = score.GradedAt,
                    SubmissionCount = logs.Count(l => l.StudentId == score.StudentId && (string.IsNullOrWhiteSpace(assignmentId) || l.AssignmentId == assignmentId))
                };
            }).ToList();

            if (!string.IsNullOrWhiteSpace(classId))
            {
                var cls = classes.FirstOrDefault(c => c.Id == classId);
                var assignment = string.IsNullOrWhiteSpace(assignmentId) ? null : assignments.FirstOrDefault(a => a.Id == assignmentId);
                var maxScore = assignment?.MaxScore ?? assignments.Where(a => a.ClassId == classId).DefaultIfEmpty().Max(a => a?.MaxScore ?? 125);
                var existingStudentIds = rows.Select(r => r.StudentId).ToHashSet();
                rows.AddRange(students
                    .Where(student => !string.IsNullOrWhiteSpace(student.Id) && student.ClassId == classId && !existingStudentIds.Contains(student.Id))
                    .Select(student => new SubmissionLeaderboardItem
                    {
                        StudentId = student.Id ?? string.Empty,
                        StudentName = FullName(student),
                        ClassId = classId,
                        ClassName = cls?.Name ?? "(Không rõ)",
                        AssignmentId = string.IsNullOrWhiteSpace(assignmentId) ? null : assignmentId,
                        AssignmentName = string.IsNullOrWhiteSpace(assignmentId) ? null : assignment?.Name,
                        ScoreValue = 0,
                        MaxScore = maxScore,
                        GradedAt = null,
                        SubmissionCount = 0
                    }));
            }

            rows = rows.OrderByDescending(r => r.ScoreValue).ThenBy(r => r.GradedAt ?? DateTime.MaxValue).ThenBy(r => r.StudentName).ToList();
            for (var i = 0; i < rows.Count; i++) rows[i].Rank = i + 1;
            return rows;
        }

        public async Task<List<SubmissionAlertResponse>> GetAlertsAsync(string portalId, bool includeDismissed = false)
        {
            EnsureObjectId(portalId, "Mã link nộp bài");
            var filter = Builders<SubmissionAlert>.Filter.Eq(a => a.PortalId, portalId);
            if (!includeDismissed) filter &= Builders<SubmissionAlert>.Filter.Eq(a => a.IsDismissed, false);
            var alerts = await _alerts.Find(filter).SortByDescending(a => a.CreatedAt).ToListAsync();
            return await MapAlertsAsync(alerts);
        }

        public async Task<int> GetUnreadAlertCountAsync(string portalId) =>
            (int)await _alerts.CountDocumentsAsync(a => a.PortalId == portalId && !a.IsRead && !a.IsDismissed);

        public async Task<bool> UpdateAlertAsync(string portalId, string alertId, bool? isRead, bool? isDismissed)
        {
            var updates = new List<UpdateDefinition<SubmissionAlert>>();
            if (isRead.HasValue) updates.Add(Builders<SubmissionAlert>.Update.Set(a => a.IsRead, isRead.Value));
            if (isDismissed.HasValue) updates.Add(Builders<SubmissionAlert>.Update.Set(a => a.IsDismissed, isDismissed.Value));
            if (updates.Count == 0) return true;
            var result = await _alerts.UpdateOneAsync(a => a.Id == alertId && a.PortalId == portalId, Builders<SubmissionAlert>.Update.Combine(updates));
            return result.ModifiedCount > 0;
        }

        public async Task<List<SubmissionLogResponse>> GetSubmissionLogsAsync(string portalId)
        {
            var logs = await _logs.Find(l => l.PortalId == portalId).SortByDescending(l => l.SubmittedAt).ToListAsync();
            var students = await _students.Find(s => s.Id != null && logs.Select(l => l.StudentId).Contains(s.Id)).ToListAsync();
            var classes = await _classes.Find(c => c.Id != null && logs.Select(l => l.ClassId).Contains(c.Id)).ToListAsync();
            var assignments = await _assignments.Find(a => logs.Select(l => l.AssignmentId).Contains(a.Id)).ToListAsync();
            return logs.Select(l => new SubmissionLogResponse
            {
                Id = l.Id,
                StudentId = l.StudentId,
                StudentName = FullName(students.FirstOrDefault(s => s.Id == l.StudentId)),
                ClassId = l.ClassId,
                ClassName = classes.FirstOrDefault(c => c.Id == l.ClassId)?.Name ?? "",
                AssignmentId = l.AssignmentId,
                AssignmentName = assignments.FirstOrDefault(a => a.Id == l.AssignmentId)?.Name ?? "",
                ScoreValue = l.ScoreValue,
                MaxScore = l.MaxScore,
                IpAddress = l.IpAddress,
                FileHash = l.FileHash,
                FileName = l.FileName,
                FileSizeBytes = l.FileSizeBytes,
                Alerts = l.Alerts,
                SubmittedAt = l.SubmittedAt
            }).ToList();
        }

        private async Task<List<string>> DetectAlertsAsync(SubmissionPortal portal, SubmissionLog current, List<SubmissionLog> previousLogs)
        {
            var tags = new List<string>();
            var now = current.SubmittedAt;
            if (!string.IsNullOrWhiteSpace(current.IpAddress))
            {
                var sameIpSubmissions = await _logs.Find(l => l.PortalId == portal.Id && l.IpAddress == current.IpAddress && l.StudentId != current.StudentId).ToListAsync();
                var studentIds = sameIpSubmissions.Select(l => l.StudentId).Append(current.StudentId).Distinct().ToList();
                if (studentIds.Count >= 2)
                {
                    tags.Add("SameIpMultipleStudents");
                    await CreateAlertAsync(portal.Id, "SameIpMultipleStudents", "High", $"IP {current.IpAddress} đã nộp bài cho {studentIds.Count} học sinh khác nhau trong link nộp bài này.", studentIds, sameIpSubmissions.Select(l => l.Id).Append(current.Id).Distinct().ToList());
                }
            }

            var recentSameStudent = await _logs.CountDocumentsAsync(l => l.PortalId == portal.Id && l.StudentId == current.StudentId && l.SubmittedAt >= now.AddMinutes(-3));
            if (recentSameStudent >= 5)
            {
                tags.Add("SpamSubmission");
                await CreateAlertAsync(portal.Id, "SpamSubmission", "Medium", "Một học sinh nộp bài quá nhiều lần trong 3 phút.", new List<string> { current.StudentId }, new List<string> { current.Id });
            }

            var previousBest = previousLogs.Where(l => l.ScoreValue.HasValue).OrderByDescending(l => l.SubmittedAt).FirstOrDefault();
            if (previousBest?.ScoreValue.HasValue == true && current.ScoreValue.HasValue && current.ScoreValue.Value - previousBest.ScoreValue.Value >= Math.Max(30, current.MaxScore * 0.5))
            {
                tags.Add("SuspiciousScoreJump");
                await CreateAlertAsync(portal.Id, "SuspiciousScoreJump", "Medium", $"Điểm tăng đột ngột từ {previousBest.ScoreValue:0.##} lên {current.ScoreValue:0.##}.", new List<string> { current.StudentId }, new List<string> { previousBest.Id, current.Id });
            }

            if (!string.IsNullOrWhiteSpace(current.FileHash))
            {
                var sameHash = await _logs.Find(l => l.PortalId == portal.Id && l.FileHash == current.FileHash && l.StudentId != current.StudentId).ToListAsync();
                if (sameHash.Count > 0)
                {
                    tags.Add("DuplicateFile");
                    await CreateAlertAsync(portal.Id, "DuplicateFile", "High", "Có nhiều học sinh nộp file giống hệt nhau.", sameHash.Select(l => l.StudentId).Append(current.StudentId).Distinct().ToList(), sameHash.Select(l => l.Id).Append(current.Id).Distinct().ToList());
                }
            }

            return tags.Distinct().ToList();
        }

        private async Task CreateAlertAsync(string portalId, string type, string severity, string message, List<string> studentIds, List<string> logIds)
        {
            await _alerts.InsertOneAsync(new SubmissionAlert { PortalId = portalId, AlertType = type, Severity = severity, Message = message, InvolvedStudentIds = studentIds, InvolvedSubmissionLogIds = logIds, CreatedAt = DateTime.UtcNow });
        }

        private async Task<bool> ShouldPersistScoreAsync(SubmissionPortal portal, string studentId, string assignmentId, double scoreValue)
        {
            if (portal.ScoringPolicy == SubmissionPortalScoringPolicies.LatestScore) return true;
            var existing = await _scoreService.GetScoreAsync(studentId, assignmentId);
            return existing?.ScoreValue == null || scoreValue >= existing.ScoreValue.Value;
        }

        private async Task ValidateSubmissionScopeAsync(SubmissionPortal portal, string classId, string studentId, string assignmentId)
        {
            if (!portal.ClassIds.Contains(classId)) throw new InvalidOperationException("Lớp không thuộc link nộp bài này.");
            if (!portal.AssignmentIds.Contains(assignmentId)) throw new InvalidOperationException("Bài tập không thuộc link nộp bài này.");
            var student = await _students.Find(s => s.Id == studentId && s.ClassId == classId && s.IsActive).FirstOrDefaultAsync();
            if (student == null) throw new InvalidOperationException("Học sinh không thuộc lớp đã chọn.");
            var assignment = await _assignments.Find(a => a.Id == assignmentId && a.ClassId == classId && a.IsActive).FirstOrDefaultAsync();
            if (assignment == null) throw new InvalidOperationException("Bài tập không thuộc lớp đã chọn.");
        }

        private async Task<SubmissionPortal> RequireOpenPortalAsync(string token, bool validateTime)
        {
            var portal = await FindActiveByTokenAsync(token) ?? throw new InvalidOperationException("Không tìm thấy link nộp bài hoặc link đã bị khóa.");
            if (validateTime)
            {
                var now = DateTime.UtcNow;
                if (portal.StartsAt.HasValue && portal.StartsAt.Value > now) throw new InvalidOperationException("Link nộp bài chưa mở.");
                if (portal.EndsAt.HasValue && portal.EndsAt.Value < now) throw new InvalidOperationException("Link nộp bài đã hết hạn.");
            }
            return portal;
        }

        private async Task<SubmissionPortal?> FindActiveByTokenAsync(string token) =>
            await _portals.Find(p => p.PublicToken == token && p.IsActive).FirstOrDefaultAsync();

        private static (string subject, string projectCode) ResolveEndpoint(string endpoint)
        {
            var normalized = endpoint.Trim().Replace("\\", "/").Trim('/').ToLowerInvariant();
            var parts = normalized.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) throw new InvalidOperationException("Endpoint chấm điểm không hợp lệ.");
            var subject = AssignmentFileSubjects.Normalize(parts[0]);
            var projectCode = string.Join('/', parts.Skip(1));
            return (subject, projectCode);
        }

        private async Task EnsureReferencesAsync(List<string> classIds, List<string> assignmentIds)
        {
            foreach (var id in classIds.Concat(assignmentIds)) EnsureObjectId(id, "Mã dữ liệu");
            var classCount = await _classes.CountDocumentsAsync(c => c.Id != null && classIds.Contains(c.Id) && c.IsActive);
            if (classCount != classIds.Count) throw new ArgumentException("Danh sách lớp có lớp không tồn tại hoặc đã khóa.");
            var assignments = await _assignments.Find(a => assignmentIds.Contains(a.Id) && a.IsActive).ToListAsync();
            if (assignments.Count != assignmentIds.Count) throw new ArgumentException("Danh sách bài tập có bài không tồn tại hoặc đã khóa.");
            if (assignments.Any(a => !classIds.Contains(a.ClassId))) throw new ArgumentException("Tất cả bài tập phải thuộc các lớp đã chọn.");
        }

        private static void ValidateRequest(CreateSubmissionPortalRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.Title)) throw new ArgumentException("Tên link nộp bài là bắt buộc.");
            if (request.ClassIds == null || request.ClassIds.Count == 0) throw new ArgumentException("Vui lòng chọn ít nhất một lớp.");
            if (request.AssignmentIds == null || request.AssignmentIds.Count == 0) throw new ArgumentException("Vui lòng chọn ít nhất một bài tập.");
            var startsAtUtc = ToUtc(request.StartsAt);
            var endsAtUtc = ToUtc(request.EndsAt);
            if (endsAtUtc.HasValue && startsAtUtc.HasValue && endsAtUtc.Value <= startsAtUtc.Value) throw new ArgumentException("Hạn nộp phải sau giờ mở nộp.");
        }

        private async Task<SubmissionPortalResponse> MapPortalAsync(SubmissionPortal portal) => new()
        {
            Id = portal.Id,
            Title = portal.Title,
            Description = portal.Description,
            PublicToken = portal.PublicToken,
            ClassIds = portal.ClassIds,
            AssignmentIds = portal.AssignmentIds,
            StartsAt = portal.StartsAt,
            EndsAt = portal.EndsAt,
            MaxSubmissionsPerStudent = portal.MaxSubmissionsPerStudent,
            ScoringPolicy = portal.ScoringPolicy,
            ShowLeaderboard = portal.ShowLeaderboard,
            ShowDetailedFeedback = portal.ShowDetailedFeedback,
            IsActive = portal.IsActive,
            CreatedAt = portal.CreatedAt,
            UnreadAlertCount = await GetUnreadAlertCountAsync(portal.Id)
        };

        private static PublicPortalAssignmentDto MapPublicAssignment(Assignment a) => new()
        {
            Id = a.Id,
            Name = a.Name,
            Description = a.Description,
            ClassId = a.ClassId,
            MaxScore = a.MaxScore,
            Subject = a.Subject,
            GradingApiEndpoint = a.GradingApiEndpoint,
            HasTemplate = !string.IsNullOrWhiteSpace(a.CurrentTemplateFileId),
            HasInstructions = !string.IsNullOrWhiteSpace(a.CurrentInstructionsFileId)
        };

        private async Task<List<SubmissionAlertResponse>> MapAlertsAsync(List<SubmissionAlert> alerts)
        {
            if (alerts.Count == 0) return new List<SubmissionAlertResponse>();

            var studentIds = alerts.SelectMany(a => a.InvolvedStudentIds).Distinct().ToList();
            var logIds = alerts.SelectMany(a => a.InvolvedSubmissionLogIds).Distinct().ToList();
            var students = studentIds.Count == 0
                ? new List<Student>()
                : await _students.Find(s => s.Id != null && studentIds.Contains(s.Id)).ToListAsync();
            var logs = logIds.Count == 0
                ? new List<SubmissionLog>()
                : await _logs.Find(l => logIds.Contains(l.Id)).ToListAsync();
            var classIds = logs.Select(l => l.ClassId)
                .Concat(students.Select(s => s.ClassId).Where(id => !string.IsNullOrWhiteSpace(id))!)
                .Distinct()
                .ToList();
            var assignmentIds = logs.Select(l => l.AssignmentId).Distinct().ToList();
            var classes = classIds.Count == 0
                ? new List<Class>()
                : await _classes.Find(c => c.Id != null && classIds.Contains(c.Id)).ToListAsync();
            var assignments = assignmentIds.Count == 0
                ? new List<Assignment>()
                : await _assignments.Find(a => assignmentIds.Contains(a.Id)).ToListAsync();

            return alerts.Select(alert => MapAlert(alert, students, logs, classes, assignments)).ToList();
        }

        private static SubmissionAlertResponse MapAlert(SubmissionAlert alert, List<Student> students, List<SubmissionLog> logs, List<Class> classes, List<Assignment> assignments) => new()
        {
            Id = alert.Id,
            AlertType = alert.AlertType,
            Severity = alert.Severity,
            Message = alert.Message,
            InvolvedStudentIds = alert.InvolvedStudentIds,
            InvolvedSubmissionLogIds = alert.InvolvedSubmissionLogIds,
            InvolvedStudents = alert.InvolvedStudentIds.Select(studentId =>
            {
                var student = students.FirstOrDefault(s => s.Id == studentId);
                var latestLog = logs
                    .Where(l => l.StudentId == studentId && alert.InvolvedSubmissionLogIds.Contains(l.Id))
                    .OrderByDescending(l => l.SubmittedAt)
                    .FirstOrDefault();
                var classId = latestLog?.ClassId ?? student?.ClassId;
                var assignmentId = latestLog?.AssignmentId;
                var cls = classes.FirstOrDefault(c => c.Id == classId);
                var assignment = assignments.FirstOrDefault(a => a.Id == assignmentId);

                return new SubmissionAlertStudentResponse
                {
                    StudentId = studentId,
                    StudentName = student == null ? "(Không rõ học sinh)" : FullName(student),
                    ClassId = classId,
                    ClassName = cls?.Name,
                    AssignmentId = assignmentId,
                    AssignmentName = assignment?.Name,
                    FileName = latestLog?.FileName,
                    IpAddress = latestLog?.IpAddress,
                    ScoreValue = latestLog?.ScoreValue,
                    MaxScore = latestLog?.MaxScore,
                    SubmittedAt = latestLog?.SubmittedAt
                };
            }).ToList(),
            IsRead = alert.IsRead,
            IsDismissed = alert.IsDismissed,
            CreatedAt = alert.CreatedAt
        };

        private static string FullName(Student? s) => s == null ? "" : $"{s.MiddleName} {s.FirstName}".Trim();
        private static DateTime? ToUtc(DateTime? value) => value.HasValue ? DateTime.SpecifyKind(value.Value, DateTimeKind.Local).ToUniversalTime() : null;
        private static void EnsureObjectId(string id, string label)
        {
            if (string.IsNullOrWhiteSpace(id) || !ObjectId.TryParse(id, out _)) throw new ArgumentException($"{label} không hợp lệ.");
        }
    }
}