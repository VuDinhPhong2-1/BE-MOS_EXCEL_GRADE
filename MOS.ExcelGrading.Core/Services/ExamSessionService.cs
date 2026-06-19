using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class ExamSessionService : IExamSessionService
    {
        private readonly IMongoCollection<ExamPublication> _examPublications;
        private readonly IMongoCollection<ExamSession> _examSessions;

        public ExamSessionService(IMongoDatabase database)
        {
            _examPublications = database.GetCollection<ExamPublication>("examPublications");
            _examSessions = database.GetCollection<ExamSession>("examSessions");
            EnsureIndexes();
        }


        public async Task<RestartCurrentProjectResponse> RestartCurrentProjectAsync(string sessionId)
        {
            EnsureObjectId(sessionId, "SessionId");

            var session = await _examSessions.Find(x => x.Id == sessionId).FirstOrDefaultAsync();
            if (session == null)
            {
                throw new ArgumentException("Không tìm thấy session.");
            }

            if (session.Status == ExamSessionStatuses.Completed)
            {
                throw new ArgumentException("Session đã hoàn thành, không thể restart project.");
            }
            if (session.IsAdvancing)
            {
                throw new ArgumentException("Session đang chuyển project, không thể restart lúc này.");
            }
            var publication = await GetPublicationOrThrow(session.PublicationId);
            var projects = publication.ProjectSequence.OrderBy(x => x.Order).ToList();

            if (session.CurrentProjectIndex < 0 || session.CurrentProjectIndex >= projects.Count)
            {
                throw new ArgumentException("CurrentProjectIndex không hợp lệ.");
            }

            var currentProject = projects[session.CurrentProjectIndex];

            var attempts = session.ProjectAttempts ?? new List<ExamSessionProjectAttempt>();

            var currentAttempts = attempts
                .Where(x => string.Equals(
                    x.ProjectCode,
                    currentProject.ProjectCode,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();

            var latestAttempt = currentAttempts
                .OrderByDescending(x => x.AttemptNo)
                .FirstOrDefault();

            if (latestAttempt != null &&
                latestAttempt.Status == ExamSessionProjectStatuses.Graded)
            {
                throw new ArgumentException("Project hiện tại đã được chấm, không thể restart.");
            }

            var nextAttemptNo = latestAttempt == null ? 1 : latestAttempt.AttemptNo + 1;

            var restartedAttempt = new ExamSessionProjectAttempt
            {
                ProjectCode = currentProject.ProjectCode,
                Subject = currentProject.Subject,
                TemplateFileName = currentProject.TemplateFileName,
                GradingApiEndpoint = currentProject.GradingApiEndpoint,
                Status = ExamSessionProjectStatuses.InProgress,
                StartedAt = DateTime.UtcNow,
                AttemptNo = nextAttemptNo
            };

            attempts.Add(restartedAttempt);

            session.ProjectAttempts = attempts;
            session.CurrentProjectStatus = ExamSessionProjectStatuses.InProgress;
            session.LastError = null;
            session.UpdatedAt = DateTime.UtcNow;

            await SaveSessionWithAdvanceConflictMappingAsync(
                session,
                "Session đang chuyển project, không thể restart lúc này.");

            return new RestartCurrentProjectResponse
            {
                State = MapState(session, publication),
                Bootstrap = MapBootstrap(currentProject)
            };
        }
        public async Task<ExamSession> StartSessionAsync(string publicationToken, StartExamSessionRequest request)
        {
            request ??= new StartExamSessionRequest();

            if (string.IsNullOrWhiteSpace(publicationToken))
            {
                throw new ArgumentException("Publication token là bắt buộc.");
            }

            if (string.IsNullOrWhiteSpace(request.StudentName))
            {
                throw new ArgumentException("Tên học sinh là bắt buộc.");
            }

            if (!string.IsNullOrWhiteSpace(request.StudentId) &&
                !ObjectId.TryParse(request.StudentId.Trim(), out _))
            {
                throw new ArgumentException("StudentId không hợp lệ.");
            }

            var publication = await _examPublications
                .Find(x => x.PublicationToken == publicationToken && x.IsActive)
                .FirstOrDefaultAsync();

            if (publication == null)
            {
                throw new ArgumentException("Không tìm thấy ca thi hoặc ca thi đã bị khóa.");
            }

            if (publication.ProjectSequence == null || publication.ProjectSequence.Count == 0)
            {
                throw new ArgumentException("Ca thi chưa có projectSequence.");
            }

            var now = DateTime.UtcNow;

            if (publication.StartsAt.HasValue && now < publication.StartsAt.Value)
            {
                throw new ArgumentException("Ca thi chưa đến giờ bắt đầu.");
            }

            if (publication.EndsAt.HasValue && now > publication.EndsAt.Value)
            {
                throw new ArgumentException("Ca thi đã hết hạn.");
            }

            if (publication.StudentIds.Any())
            {
                if (string.IsNullOrWhiteSpace(request.StudentId))
                {
                    throw new ArgumentException("StudentId là bắt buộc cho ca thi này.");
                }

                if (!publication.StudentIds.Contains(request.StudentId.Trim(), StringComparer.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("Học sinh không nằm trong danh sách được phép thi.");
                }
            }

            var normalizedStudentId = string.IsNullOrWhiteSpace(request.StudentId)
                ? null
                : request.StudentId.Trim();
            var normalizedStudentName = request.StudentName.Trim();

            if (!string.IsNullOrWhiteSpace(normalizedStudentId))
            {
                var existingSession = await _examSessions
                    .Find(x =>
                        x.PublicationId == publication.Id &&
                        x.StudentId == normalizedStudentId &&
                        x.Status == ExamSessionStatuses.InProgress)
                    .SortByDescending(x => x.StartedAt)
                    .FirstOrDefaultAsync();

                if (existingSession != null)
                {
                    if (!string.Equals(existingSession.StudentName, normalizedStudentName, StringComparison.Ordinal))
                    {
                        existingSession.StudentName = normalizedStudentName;
                        existingSession.UpdatedAt = now;
                        await ReplaceSessionWithOptimisticConcurrencyAsync(existingSession);
                    }

                    return existingSession;
                }
            }

            var firstProject = publication.ProjectSequence.OrderBy(x => x.Order).First();

            var session = new ExamSession
            {
                PublicationId = publication.Id,
                PublicationToken = publication.PublicationToken,
                StudentId = normalizedStudentId,
                StudentName = normalizedStudentName,
                CurrentProjectIndex = 0,
                CurrentProjectStatus = ExamSessionProjectStatuses.InProgress,
                Status = ExamSessionStatuses.InProgress,
                ProjectAttempts = new List<ExamSessionProjectAttempt>
                {
                    CreateAttempt(firstProject, ExamSessionProjectStatuses.InProgress)
                },
                StartedAt = now,
                CreatedAt = now
            };

            await _examSessions.InsertOneAsync(session);
            return session;
        }

        private void EnsureIndexes()
        {
            var publicationStudentStatusIndex = new CreateIndexModel<ExamSession>(
                Builders<ExamSession>.IndexKeys
                    .Ascending(x => x.PublicationId)
                    .Ascending(x => x.StudentId)
                    .Ascending(x => x.Status)
                    .Descending(x => x.StartedAt));

            _examSessions.Indexes.CreateOne(publicationStudentStatusIndex);
        }

        public async Task<ExamSessionStateDto?> GetStateAsync(string publicationToken, string sessionId)
        {
            EnsureObjectId(sessionId, "SessionId");

            var session = await _examSessions.Find(x =>
                    x.Id == sessionId &&
                    x.PublicationToken == publicationToken)
                .FirstOrDefaultAsync();

            if (session == null)
            {
                return null;
            }

            var publication = await GetPublicationOrThrow(session.PublicationId);
            return MapState(session, publication);
        }

        public async Task<ExamSessionProjectBootstrapDto?> GetCurrentProjectBootstrapAsync(string sessionId)
        {
            EnsureObjectId(sessionId, "SessionId");

            var session = await _examSessions.Find(x => x.Id == sessionId).FirstOrDefaultAsync();
            if (session == null || session.Status == ExamSessionStatuses.Completed)
            {
                return null;
            }

            var publication = await GetPublicationOrThrow(session.PublicationId);
            var projects = publication.ProjectSequence.OrderBy(x => x.Order).ToList();

            if (session.CurrentProjectIndex < 0 || session.CurrentProjectIndex >= projects.Count)
            {
                return null;
            }

            return MapBootstrap(projects[session.CurrentProjectIndex]);
        }

        public async Task<ExamSessionProjectBootstrapDto?> GetCurrentProjectBootstrapAsync(
            string publicationToken,
            string sessionId)
        {
            await EnsureSessionBelongsToPublicationAsync(publicationToken, sessionId);
            return await GetCurrentProjectBootstrapAsync(sessionId);
        }

        public async Task UploadScoreAsync(string sessionId, string projectCode, ScoreUploadRequest request)
        {
            request ??= new ScoreUploadRequest();
            EnsureObjectId(sessionId, "SessionId");

            if (string.IsNullOrWhiteSpace(projectCode))
            {
                throw new ArgumentException("ProjectCode là bắt buộc.");
            }

            var session = await _examSessions.Find(x => x.Id == sessionId).FirstOrDefaultAsync();
            if (session == null)
            {
                throw new ArgumentException("Không tìm thấy session.");
            }

            if (session.Status == ExamSessionStatuses.Completed)
            {
                throw new ArgumentException("Session đã hoàn thành.");
            }
            if (session.IsAdvancing)
            {
                throw new ArgumentException("Session đang chuyển project, không thể upload điểm lúc này.");
            }
            var publication = await GetPublicationOrThrow(session.PublicationId);
            var projects = publication.ProjectSequence.OrderBy(x => x.Order).ToList();

            if (session.CurrentProjectIndex < 0 || session.CurrentProjectIndex >= projects.Count)
            {
                throw new ArgumentException("CurrentProjectIndex không hợp lệ.");
            }

            var currentProject = projects[session.CurrentProjectIndex];

            if (!string.Equals(currentProject.ProjectCode, projectCode.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("ProjectCode không khớp với project hiện tại.");
            }

            var attempts = session.ProjectAttempts ?? new List<ExamSessionProjectAttempt>();
            var attempt = attempts.LastOrDefault(x =>
                string.Equals(x.ProjectCode, currentProject.ProjectCode, StringComparison.OrdinalIgnoreCase));

            if (attempt == null)
            {
                attempt = CreateAttempt(currentProject, ExamSessionProjectStatuses.InProgress);
                attempts.Add(attempt);
            }

            var now = DateTime.UtcNow;

            attempt.WorkingFilePath = NormalizeNullable(request.WorkingFilePath);
            attempt.SubmittedAt ??= now;

            if (!request.IsSuccess)
            {
                attempt.Status = ExamSessionProjectStatuses.Error;
                attempt.Feedback = NormalizeNullable(request.ErrorMessage) ?? "Chấm điểm thất bại.";
                session.CurrentProjectStatus = ExamSessionProjectStatuses.Error;
                session.LastError = attempt.Feedback;
            }
            else
            {
                attempt.Status = ExamSessionProjectStatuses.Graded;
                attempt.GradedAt = now;
                attempt.Score = request.Score;
                attempt.MaxScore = request.MaxScore;
                attempt.Feedback = NormalizeNullable(request.Feedback);

                session.CurrentProjectStatus = ExamSessionProjectStatuses.Graded;
                session.LastError = null;
            }

            session.ProjectAttempts = attempts;
            session.CompletedProjectCount = attempts.Count(x => x.Status == ExamSessionProjectStatuses.Graded);
            session.AggregateScore = CalculateAggregateScore(attempts);
            session.UpdatedAt = now;

            await SaveSessionWithAdvanceConflictMappingAsync(
                session,
                "Session đang chuyển project, không thể upload điểm lúc này.");
        }

        public async Task<AdvanceExamSessionResponse> AdvanceAsync(string sessionId)
        {
            EnsureObjectId(sessionId, "SessionId");

            var now = DateTime.UtcNow;

            var lockFilter = Builders<ExamSession>.Filter.And(
                Builders<ExamSession>.Filter.Eq(x => x.Id, sessionId),
                Builders<ExamSession>.Filter.Eq(x => x.Status, ExamSessionStatuses.InProgress),
                Builders<ExamSession>.Filter.Eq(x => x.IsAdvancing, false)
            );

            var lockUpdate = Builders<ExamSession>.Update
                .Set(x => x.IsAdvancing, true)
                .Set(x => x.AdvanceStartedAt, now)
                .Set(x => x.UpdatedAt, now)
                .Inc(x => x.Version, 1);

            var lockedSession = await _examSessions.FindOneAndUpdateAsync(
                lockFilter,
                lockUpdate,
                new FindOneAndUpdateOptions<ExamSession>
                {
                    ReturnDocument = ReturnDocument.After
                });

            if (lockedSession == null)
            {
                var existingSession = await _examSessions
                    .Find(x => x.Id == sessionId)
                    .FirstOrDefaultAsync();

                if (existingSession == null)
                {
                    throw new ArgumentException("Không tìm thấy session.");
                }

                if (existingSession.Status == ExamSessionStatuses.Completed)
                {
                    var completedPublication = await GetPublicationOrThrow(existingSession.PublicationId);

                    return new AdvanceExamSessionResponse
                    {
                        IsCompleted = true,
                        State = MapState(existingSession, completedPublication)
                    };
                }

                if (existingSession.IsAdvancing)
                {
                    throw new ArgumentException("Session đang chuyển project, vui lòng chờ.");
                }

                throw new ArgumentException("Không thể khóa session để chuyển project.");
            }

            try
            {
                var publication = await GetPublicationOrThrow(lockedSession.PublicationId);
                var projects = publication.ProjectSequence.OrderBy(x => x.Order).ToList();

                if (lockedSession.CurrentProjectIndex < 0 ||
                    lockedSession.CurrentProjectIndex >= projects.Count)
                {
                    throw new ArgumentException("CurrentProjectIndex không hợp lệ.");
                }

                var currentProject = projects[lockedSession.CurrentProjectIndex];

                var currentAttempt = lockedSession.ProjectAttempts.LastOrDefault(x =>
                    string.Equals(
                        x.ProjectCode,
                        currentProject.ProjectCode,
                        StringComparison.OrdinalIgnoreCase));

                if (currentAttempt == null ||
                    currentAttempt.Status != ExamSessionProjectStatuses.Graded)
                {
                    throw new ArgumentException("Project hiện tại chưa được chấm thành công, không thể chuyển project.");
                }

                var nextIndex = lockedSession.CurrentProjectIndex + 1;

                if (nextIndex >= projects.Count)
                {
                    lockedSession.Status = ExamSessionStatuses.Completed;
                    lockedSession.CurrentProjectStatus = ExamSessionProjectStatuses.Graded;
                    lockedSession.CompletedAt = DateTime.UtcNow;
                    lockedSession.CompletedProjectCount = lockedSession.ProjectAttempts
                        .Count(x => x.Status == ExamSessionProjectStatuses.Graded);
                    lockedSession.AggregateScore = CalculateAggregateScore(lockedSession.ProjectAttempts);
                    lockedSession.IsAdvancing = false;
                    lockedSession.AdvanceStartedAt = null;
                    lockedSession.UpdatedAt = DateTime.UtcNow;

                    await ReplaceSessionWithOptimisticConcurrencyAsync(lockedSession);

                    return new AdvanceExamSessionResponse
                    {
                        IsCompleted = true,
                        State = MapState(lockedSession, publication)
                    };
                }

                var nextProject = projects[nextIndex];

                lockedSession.CurrentProjectIndex = nextIndex;
                lockedSession.CurrentProjectStatus = ExamSessionProjectStatuses.InProgress;
                lockedSession.LastError = null;
                lockedSession.IsAdvancing = false;
                lockedSession.AdvanceStartedAt = null;
                lockedSession.UpdatedAt = DateTime.UtcNow;

                var existingNextAttempt = lockedSession.ProjectAttempts.LastOrDefault(x =>
                    string.Equals(
                        x.ProjectCode,
                        nextProject.ProjectCode,
                        StringComparison.OrdinalIgnoreCase));

                if (existingNextAttempt == null)
                {
                    lockedSession.ProjectAttempts.Add(CreateAttempt(
                        nextProject,
                        ExamSessionProjectStatuses.InProgress));
                }

                lockedSession.CompletedProjectCount = lockedSession.ProjectAttempts
                    .Count(x => x.Status == ExamSessionProjectStatuses.Graded);

                lockedSession.AggregateScore = CalculateAggregateScore(lockedSession.ProjectAttempts);

                await ReplaceSessionWithOptimisticConcurrencyAsync(lockedSession);

                return new AdvanceExamSessionResponse
                {
                    IsCompleted = false,
                    State = MapState(lockedSession, publication)
                };
            }
            catch
            {
                var unlockUpdate = Builders<ExamSession>.Update
                    .Set(x => x.IsAdvancing, false)
                    .Set(x => x.AdvanceStartedAt, null)
                    .Set(x => x.UpdatedAt, DateTime.UtcNow)
                    .Inc(x => x.Version, 1);

                await _examSessions.UpdateOneAsync(
                    x => x.Id == sessionId,
                    unlockUpdate);

                throw;
            }
        }

        public async Task UploadScoreAsync(
            string publicationToken,
            string sessionId,
            string projectCode,
            ScoreUploadRequest request)
        {
            await EnsureSessionBelongsToPublicationAsync(publicationToken, sessionId);
            await UploadScoreAsync(sessionId, projectCode, request);
        }

        public async Task<AdvanceExamSessionResponse> AdvanceAsync(string publicationToken, string sessionId)
        {
            await EnsureSessionBelongsToPublicationAsync(publicationToken, sessionId);
            return await AdvanceAsync(sessionId);
        }

        public async Task<RestartCurrentProjectResponse> RestartCurrentProjectAsync(
            string publicationToken,
            string sessionId)
        {
            await EnsureSessionBelongsToPublicationAsync(publicationToken, sessionId);
            return await RestartCurrentProjectAsync(sessionId);
        }

        private async Task<ExamPublication> GetPublicationOrThrow(string publicationId)
        {
            var publication = await _examPublications
                .Find(x => x.Id == publicationId)
                .FirstOrDefaultAsync();

            if (publication == null)
            {
                throw new ArgumentException("Không tìm thấy publication.");
            }

            return publication;
        }

        private static ExamSessionProjectAttempt CreateAttempt(
            ExamPublicationProject project,
            string status)
        {
            return new ExamSessionProjectAttempt
            {
                ProjectCode = project.ProjectCode,
                Subject = project.Subject,
                TemplateFileName = project.TemplateFileName,
                GradingApiEndpoint = project.GradingApiEndpoint,
                Status = status,
                StartedAt = DateTime.UtcNow,
                AttemptNo = 1
            };
        }

        private static ExamSessionStateDto MapState(ExamSession session, ExamPublication publication)
        {
            var projects = publication.ProjectSequence.OrderBy(x => x.Order).ToList();

            ExamPublicationProject? currentProject = null;
            ExamPublicationProject? nextProject = null;

            if (session.CurrentProjectIndex >= 0 && session.CurrentProjectIndex < projects.Count)
            {
                currentProject = projects[session.CurrentProjectIndex];
            }

            var nextIndex = session.CurrentProjectIndex + 1;
            if (nextIndex >= 0 && nextIndex < projects.Count)
            {
                nextProject = projects[nextIndex];
            }

            return new ExamSessionStateDto
            {
                SessionId = session.Id,
                PublicationId = publication.Id,
                PublicationName = publication.Name,
                CurrentProjectIndex = session.CurrentProjectIndex,
                CurrentProjectNumber = session.CurrentProjectIndex + 1,
                TotalProjectCount = projects.Count,
                CompletedProjectCount = session.CompletedProjectCount,
                Status = session.Status,
                CurrentProjectStatus = session.CurrentProjectStatus,
                IsAdvancing = session.IsAdvancing,
                AdvanceStartedAt = session.AdvanceStartedAt,
                CompletedAt = session.CompletedAt,
                CurrentProject = currentProject == null ? null : MapBootstrap(currentProject),
                NextProject = nextProject == null ? null : MapBootstrap(nextProject),
                AggregateScore = session.AggregateScore,
                LastError = session.LastError
            };
        }

        private static ExamSessionProjectBootstrapDto MapBootstrap(ExamPublicationProject project)
        {
            return new ExamSessionProjectBootstrapDto
            {
                Order = project.Order,
                ProjectCode = project.ProjectCode,
                Subject = project.Subject,
                TemplateFileName = project.TemplateFileName,
                InstructionsFileName = project.InstructionsFileName,
                InstructionsText = project.InstructionsText,
                HelpFileName = project.HelpFileName,
                HelpText = project.HelpText,
                GradingApiEndpoint = project.GradingApiEndpoint,
                TaskSnapshot = project.TaskSnapshot.Select(x => new ExamPublicationTaskSnapshotItemDto
                {
                    TaskId = x.TaskId,
                    TaskName = x.TaskName,
                    MaxScore = x.MaxScore,
                    Instructions = x.Instructions
                }).ToList(),
                ModeRules = project.ModeRules == null
                    ? null
                    : new ExamPublicationModeRulesDto
                    {
                        Mode = project.ModeRules.Mode,
                        ShowFeedback = project.ModeRules.ShowFeedback,
                        AllowRestart = project.ModeRules.AllowRestart,
                        AllowNextProject = project.ModeRules.AllowNextProject
                    }
            };
        }

        private static double? CalculateAggregateScore(List<ExamSessionProjectAttempt> attempts)
        {
            var gradedScores = attempts
                .Where(x => x.Status == ExamSessionProjectStatuses.Graded && x.Score.HasValue)
                .Select(x => x.Score!.Value)
                .ToList();

            if (!gradedScores.Any())
            {
                return null;
            }

            return gradedScores.Sum();
        }

        private async Task EnsureSessionBelongsToPublicationAsync(string publicationToken, string sessionId)
        {
            if (string.IsNullOrWhiteSpace(publicationToken))
            {
                throw new ArgumentException("Publication token lÃ  báº¯t buá»™c.");
            }

            EnsureObjectId(sessionId, "SessionId");

            var sessionExists = await _examSessions
                .Find(x => x.Id == sessionId && x.PublicationToken == publicationToken)
                .AnyAsync();

            if (!sessionExists)
            {
                throw new ArgumentException("KhÃ´ng tÃ¬m tháº¥y phiÃªn thi phÃ¹ há»£p vá»›i publication token.");
            }
        }

        private async Task ReplaceSessionWithOptimisticConcurrencyAsync(ExamSession session)
        {
            var expectedVersion = session.Version;
            session.Version = expectedVersion + 1;

            var result = await _examSessions.ReplaceOneAsync(
                x => x.Id == session.Id && x.Version == expectedVersion,
                session);

            if (result.ModifiedCount == 0)
            {
                throw new InvalidOperationException("Session đã thay đổi bởi request khác. Vui lòng thử lại.");
            }
        }

        private async Task SaveSessionWithAdvanceConflictMappingAsync(
            ExamSession session,
            string advancingMessage)
        {
            try
            {
                await ReplaceSessionWithOptimisticConcurrencyAsync(session);
            }
            catch (InvalidOperationException)
            {
                var latestSession = await _examSessions
                    .Find(x => x.Id == session.Id)
                    .FirstOrDefaultAsync();

                if (latestSession?.IsAdvancing == true)
                {
                    throw new ArgumentException(advancingMessage);
                }

                if (latestSession?.Status == ExamSessionStatuses.Completed)
                {
                    throw new ArgumentException("Session đã hoàn thành.");
                }

                throw;
            }
        }

        private static void EnsureObjectId(string? value, string fieldLabel)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new ArgumentException($"{fieldLabel} không hợp lệ.");
            }
        }

        private static string? NormalizeNullable(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }
}
