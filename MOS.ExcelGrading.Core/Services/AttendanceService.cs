using Microsoft.Extensions.Logging;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class AttendanceService : IAttendanceService
    {
        private readonly IMongoCollection<StudentScheduleAttendance> _attendances;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<Class> _classes;
        private readonly IMongoCollection<TeacherSchedule> _schedules;
        private readonly ILogger<AttendanceService> _logger;
        private static int _indexInitialized;

        public AttendanceService(IMongoDatabase database, ILogger<AttendanceService> logger)
        {
            _attendances = database.GetCollection<StudentScheduleAttendance>("scheduleAttendances");
            _students = database.GetCollection<Student>("students");
            _classes = database.GetCollection<Class>("Classes");
            _schedules = database.GetCollection<TeacherSchedule>("teacherSchedules");
            _logger = logger;

            if (Interlocked.Exchange(ref _indexInitialized, 1) == 0)
            {
                var uniqueIndex = new CreateIndexModel<StudentScheduleAttendance>(
                    Builders<StudentScheduleAttendance>.IndexKeys
                        .Ascending(x => x.OwnerId)
                        .Ascending(x => x.ScheduleId)
                        .Ascending(x => x.StudentId),
                    new CreateIndexOptions { Unique = true });

                var weekIndex = new CreateIndexModel<StudentScheduleAttendance>(
                    Builders<StudentScheduleAttendance>.IndexKeys
                        .Ascending(x => x.OwnerId)
                        .Ascending(x => x.Date)
                        .Ascending(x => x.ClassId));

                _attendances.Indexes.CreateMany(new[] { uniqueIndex, weekIndex });
            }
        }

        public async Task<ScheduleAttendanceResponse> GetScheduleAttendanceAsync(TeacherSchedule schedule, string ownerId, bool isAdmin)
        {
            var classInfo = await ResolveClassAsync(schedule, ownerId, isAdmin);
            if (classInfo == null)
            {
                throw new InvalidOperationException("Lich day chua gan lop hop le, khong the diem danh.");
            }
            if (string.IsNullOrWhiteSpace(classInfo.Id))
            {
                throw new InvalidOperationException("Khong tim thay ma lop hop le de diem danh.");
            }

            var students = await _students
                .Find(s => s.ClassId == classInfo.Id && s.IsActive)
                .SortBy(s => s.MiddleName)
                .ThenBy(s => s.FirstName)
                .ToListAsync();

            var attendanceRecords = await _attendances
                .Find(x => x.OwnerId == ownerId && x.ScheduleId == (schedule.Id ?? string.Empty))
                .ToListAsync();

            var attendanceByStudent = attendanceRecords.ToDictionary(x => x.StudentId, x => x);
            var rows = new List<ScheduleAttendanceStudentResponse>(students.Count);

            foreach (var student in students)
            {
                StudentScheduleAttendance? att = null;
                var hasAttendance = !string.IsNullOrWhiteSpace(student.Id)
                    && attendanceByStudent.TryGetValue(student.Id, out att);

                var attendanceStatus = hasAttendance
                    ? NormalizeStatus(att?.Status)
                    : AttendanceStatus.Present;

                rows.Add(new ScheduleAttendanceStudentResponse
                {
                    StudentId = student.Id ?? string.Empty,
                    MiddleName = student.MiddleName,
                    FirstName = student.FirstName,
                    FullName = $"{student.MiddleName} {student.FirstName}".Trim(),
                    StudentStatus = student.Status,
                    AttendanceStatus = attendanceStatus,
                    Note = hasAttendance ? att?.Note : null,
                    MarkedAt = hasAttendance ? att?.MarkedAt : null
                });
            }

            var sameRoomSessionSchedules = await GetSameRoomSessionSchedulesAsync(schedule, ownerId);
            var roomSessionContext = await BuildRoomSessionContextAsync(schedule, classInfo, sameRoomSessionSchedules);
            var reports = BuildReportsResponse(schedule, roomSessionContext);

            return BuildResponse(schedule, classInfo.Id, rows, reports, roomSessionContext);
        }

        public async Task<ScheduleAttendanceResponse> SaveScheduleAttendanceAsync(
            TeacherSchedule schedule,
            string ownerId,
            bool isAdmin,
            IReadOnlyCollection<SaveScheduleAttendanceItem> items,
            ScheduleReportsRequest? reports,
            string updatedBy)
        {
            var classInfo = await ResolveClassAsync(schedule, ownerId, isAdmin);
            if (classInfo == null)
            {
                throw new InvalidOperationException("Lich day chua gan lop hop le, khong the luu diem danh.");
            }
            if (string.IsNullOrWhiteSpace(classInfo.Id))
            {
                throw new InvalidOperationException("Khong tim thay ma lop hop le de luu diem danh.");
            }

            var classStudentIds = await _students
                .Find(s => s.ClassId == classInfo.Id && s.IsActive)
                .Project(s => s.Id)
                .ToListAsync();

            var validStudentIds = new HashSet<string>(
                classStudentIds.Where(id => !string.IsNullOrWhiteSpace(id))!);

            var normalizedItems = items
                .Where(i => !string.IsNullOrWhiteSpace(i.StudentId))
                .GroupBy(i => i.StudentId.Trim())
                .Select(g => g.Last())
                .ToList();

            var operations = new List<WriteModel<StudentScheduleAttendance>>(normalizedItems.Count);
            var now = DateTime.UtcNow;

            foreach (var item in normalizedItems)
            {
                var studentId = item.StudentId.Trim();
                if (!validStudentIds.Contains(studentId))
                {
                    continue;
                }

                var normalizedStatus = NormalizeStatus(item.Status);
                var filter = Builders<StudentScheduleAttendance>.Filter.And(
                    Builders<StudentScheduleAttendance>.Filter.Eq(x => x.OwnerId, ownerId),
                    Builders<StudentScheduleAttendance>.Filter.Eq(x => x.ScheduleId, schedule.Id),
                    Builders<StudentScheduleAttendance>.Filter.Eq(x => x.StudentId, studentId)
                );

                if (normalizedStatus == AttendanceStatus.Present)
                {
                    operations.Add(new DeleteOneModel<StudentScheduleAttendance>(filter));
                    continue;
                }

                var update = Builders<StudentScheduleAttendance>.Update
                    .Set(x => x.OwnerId, ownerId)
                    .Set(x => x.ScheduleId, schedule.Id)
                    .Set(x => x.SchoolId, schedule.SchoolId)
                    .Set(x => x.ClassId, classInfo.Id)
                    .Set(x => x.StudentId, studentId)
                    .Set(x => x.Date, schedule.Date)
                    .Set(x => x.Status, normalizedStatus)
                    .Set(x => x.Note, item.Note?.Trim())
                    .Set(x => x.MarkedAt, now)
                    .Set(x => x.MarkedBy, updatedBy)
                    .Set(x => x.UpdatedAt, now)
                    .SetOnInsert(x => x.CreatedAt, now);

                operations.Add(new UpdateOneModel<StudentScheduleAttendance>(filter, update) { IsUpsert = true });
            }

            if (operations.Count > 0)
            {
                await _attendances.BulkWriteAsync(operations, new BulkWriteOptions { IsOrdered = false });
            }

            if (reports != null && !string.IsNullOrWhiteSpace(schedule.Id))
            {
                var sameRoomSessionSchedules = await GetSameRoomSessionSchedulesAsync(schedule, ownerId);
                var roomSessionContext = await BuildRoomSessionContextAsync(schedule, classInfo, sameRoomSessionSchedules);
                var normalizedBundle = MapReportsRequestToModel(reports, schedule, roomSessionContext);
                await SaveReportsAsync(schedule, ownerId, updatedBy, normalizedBundle, sameRoomSessionSchedules, roomSessionContext);
            }

            var refreshedSchedule = await ReloadScheduleAsync(schedule, ownerId);
            return await GetScheduleAttendanceAsync(refreshedSchedule, ownerId, isAdmin);
        }

        private async Task SaveReportsAsync(
            TeacherSchedule schedule,
            string ownerId,
            string updatedBy,
            ScheduleReportBundle reports,
            IReadOnlyCollection<TeacherSchedule> sameRoomSessionSchedules,
            ScheduleRoomSessionContextResponse roomSessionContext)
        {
            if (string.IsNullOrWhiteSpace(schedule.Id))
            {
                return;
            }

            var ownerForFilter = GetOwnerForSchedule(schedule, ownerId);
            var now = DateTime.UtcNow;

            var currentFilter = Builders<TeacherSchedule>.Filter.And(
                Builders<TeacherSchedule>.Filter.Eq(x => x.Id, schedule.Id),
                Builders<TeacherSchedule>.Filter.Eq(x => x.OwnerId, ownerForFilter)
            );

            var fullUpdate = Builders<TeacherSchedule>.Update
                .Set(x => x.Reports, reports)
                .Set(x => x.UpdatedAt, now)
                .Set(x => x.UpdatedBy, updatedBy);

            await _schedules.UpdateOneAsync(currentFilter, fullUpdate);

            if (!roomSessionContext.IsSharedRoomSession)
            {
                return;
            }

            var otherIds = sameRoomSessionSchedules
                .Select(x => x.Id)
                .Where(id => !string.IsNullOrWhiteSpace(id) && !string.Equals(id, schedule.Id, StringComparison.Ordinal))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToList();

            if (otherIds.Count == 0)
            {
                return;
            }

            var sharedFilter = Builders<TeacherSchedule>.Filter.And(
                Builders<TeacherSchedule>.Filter.In(x => x.Id, otherIds),
                Builders<TeacherSchedule>.Filter.Eq(x => x.OwnerId, ownerForFilter)
            );

            var sharedUpdate = Builders<TeacherSchedule>.Update
                .Set(x => x.Reports.EndLesson, reports.EndLesson)
                .Set(x => x.UpdatedAt, now)
                .Set(x => x.UpdatedBy, updatedBy);

            await _schedules.UpdateManyAsync(sharedFilter, sharedUpdate);
        }

        private async Task<TeacherSchedule> ReloadScheduleAsync(TeacherSchedule schedule, string ownerId)
        {
            if (string.IsNullOrWhiteSpace(schedule.Id))
            {
                return schedule;
            }

            var ownerForFilter = GetOwnerForSchedule(schedule, ownerId);
            var refreshed = await _schedules
                .Find(x => x.Id == schedule.Id && x.OwnerId == ownerForFilter)
                .FirstOrDefaultAsync();

            return refreshed ?? schedule;
        }

        private async Task<List<TeacherSchedule>> GetSameRoomSessionSchedulesAsync(TeacherSchedule schedule, string ownerId)
        {
            var ownerForFilter = GetOwnerForSchedule(schedule, ownerId);
            var sessionCode = ResolveSessionCode(schedule.StartTime);
            var normalizedRoom = NormalizeKey(schedule.RoomName);

            if (string.IsNullOrWhiteSpace(normalizedRoom))
            {
                return new List<TeacherSchedule> { schedule };
            }

            var filter = Builders<TeacherSchedule>.Filter.And(
                Builders<TeacherSchedule>.Filter.Eq(x => x.OwnerId, ownerForFilter),
                Builders<TeacherSchedule>.Filter.Eq(x => x.Date, schedule.Date),
                Builders<TeacherSchedule>.Filter.Eq(x => x.IsActive, true)
            );

            if (!string.IsNullOrWhiteSpace(schedule.SchoolId))
            {
                filter &= Builders<TeacherSchedule>.Filter.Eq(x => x.SchoolId, schedule.SchoolId);
            }

            var sameDaySchedules = await _schedules.Find(filter).ToListAsync();

            var sameRoomSession = sameDaySchedules
                .Where(x =>
                    NormalizeKey(x.RoomName) == normalizedRoom &&
                    ResolveSessionCode(x.StartTime) == sessionCode)
                .ToList();

            if (sameRoomSession.Count == 0)
            {
                return new List<TeacherSchedule> { schedule };
            }

            var currentExists = !string.IsNullOrWhiteSpace(schedule.Id)
                && sameRoomSession.Any(x => string.Equals(x.Id, schedule.Id, StringComparison.Ordinal));

            if (!currentExists)
            {
                sameRoomSession.Add(schedule);
            }

            return sameRoomSession
                .GroupBy(x => x.Id ?? $"{x.ClassId}|{x.ClassName}|{x.StartTime}|{x.EndTime}")
                .Select(g => g.First())
                .ToList();
        }

        private async Task<ScheduleRoomSessionContextResponse> BuildRoomSessionContextAsync(
            TeacherSchedule schedule,
            Class classInfo,
            IReadOnlyCollection<TeacherSchedule> sameRoomSessionSchedules)
        {
            var schedules = sameRoomSessionSchedules.Count == 0
                ? new List<TeacherSchedule> { schedule }
                : sameRoomSessionSchedules.ToList();

            var uniqueClassRecords = schedules
                .Select(x => new
                {
                    ClassId = string.IsNullOrWhiteSpace(x.ClassId) ? null : x.ClassId!.Trim(),
                    ClassName = string.IsNullOrWhiteSpace(x.ClassName) ? classInfo.Name : x.ClassName.Trim(),
                })
                .GroupBy(x => x.ClassId ?? x.ClassName, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            var classIds = uniqueClassRecords
                .Where(x => !string.IsNullOrWhiteSpace(x.ClassId))
                .Select(x => x.ClassId!)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            var classMap = classIds.Count > 0
                ? (await _classes.Find(Builders<Class>.Filter.In(x => x.Id, classIds)).ToListAsync())
                    .Where(x => !string.IsNullOrWhiteSpace(x.Id))
                    .ToDictionary(x => x.Id!, x => x)
                : new Dictionary<string, Class>();

            var classSummaries = new List<ScheduleRoomClassSummaryResponse>();
            foreach (var item in uniqueClassRecords)
            {
                if (!string.IsNullOrWhiteSpace(item.ClassId) && classMap.TryGetValue(item.ClassId, out var classEntity))
                {
                    classSummaries.Add(new ScheduleRoomClassSummaryResponse
                    {
                        ClassId = item.ClassId,
                        ClassName = string.IsNullOrWhiteSpace(item.ClassName) ? classEntity.Name : item.ClassName,
                        CurrentStudents = classEntity.CurrentStudents,
                        MaxStudents = classEntity.MaxStudents
                    });
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.ClassId) && item.ClassId == classInfo.Id)
                {
                    classSummaries.Add(new ScheduleRoomClassSummaryResponse
                    {
                        ClassId = classInfo.Id,
                        ClassName = string.IsNullOrWhiteSpace(item.ClassName) ? classInfo.Name : item.ClassName,
                        CurrentStudents = classInfo.CurrentStudents,
                        MaxStudents = classInfo.MaxStudents
                    });
                    continue;
                }

                classSummaries.Add(new ScheduleRoomClassSummaryResponse
                {
                    ClassId = item.ClassId,
                    ClassName = string.IsNullOrWhiteSpace(item.ClassName) ? schedule.ClassName : item.ClassName,
                    CurrentStudents = 0,
                    MaxStudents = null
                });
            }

            classSummaries = classSummaries
                .OrderBy(x => x.ClassName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new ScheduleRoomSessionContextResponse
            {
                SessionLabel = ResolveSessionLabel(schedule.StartTime),
                IsSharedRoomSession = classSummaries.Count > 1,
                SharedClasses = classSummaries,
                SharedClassStudentSummary = BuildClassStudentSummaryText(classSummaries)
            };
        }

        private static ScheduleReportsResponse BuildReportsResponse(
            TeacherSchedule schedule,
            ScheduleRoomSessionContextResponse roomSessionContext)
        {
            var persisted = schedule.Reports ?? new ScheduleReportBundle();
            persisted.StartLesson ??= new StartLessonReport();
            persisted.Professional ??= new ProfessionalReport();
            persisted.EndLesson ??= new EndLessonReport();

            return new ScheduleReportsResponse
            {
                StartLesson = new StartLessonReportResponse
                {
                    TeacherName = persisted.StartLesson.TeacherName,
                    AssistantName = persisted.StartLesson.AssistantName,
                    RoomName = FirstNonEmpty(persisted.StartLesson.RoomName, schedule.RoomName),
                    TotalMachines = persisted.StartLesson.TotalMachines,
                    BrokenMachinesSummary = persisted.StartLesson.BrokenMachinesSummary,
                    MissingMachinesForStudents = persisted.StartLesson.MissingMachinesForStudents,
                    NetSupportStatus = persisted.StartLesson.NetSupportStatus,
                    AudioStatus = persisted.StartLesson.AudioStatus,
                    CoolingStatus = persisted.StartLesson.CoolingStatus,
                    HygieneStatus = persisted.StartLesson.HygieneStatus
                },
                Professional = new ProfessionalReportResponse
                {
                    TeacherName = persisted.Professional.TeacherName,
                    ClassName = FirstNonEmpty(persisted.Professional.ClassName, schedule.ClassName),
                    SubjectName = FirstNonEmpty(persisted.Professional.SubjectName, schedule.Subject),
                    TeachingMaterials = persisted.Professional.TeachingMaterials,
                    TeachingContent = persisted.Professional.TeachingContent,
                    PlannedLessons = persisted.Professional.PlannedLessons,
                    TaughtLessons = persisted.Professional.TaughtLessons,
                    OngoingPracticeCompletions = persisted.Professional.OngoingPracticeCompletions,
                    GmetrixResultRate = persisted.Professional.GmetrixResultRate
                },
                EndLesson = new EndLessonReportResponse
                {
                    TeacherName = persisted.EndLesson.TeacherName,
                    AssistantName = persisted.EndLesson.AssistantName,
                    RoomName = FirstNonEmpty(persisted.EndLesson.RoomName, schedule.RoomName),
                    TotalMachines = persisted.EndLesson.TotalMachines,
                    ClassStudentCountSummary = FirstNonEmpty(
                        persisted.EndLesson.ClassStudentCountSummary,
                        roomSessionContext.SharedClassStudentSummary),
                    StudentMaterialCoverageRate = persisted.EndLesson.StudentMaterialCoverageRate,
                    BrokenMachinesSummary = persisted.EndLesson.BrokenMachinesSummary,
                    NetSupportStatus = persisted.EndLesson.NetSupportStatus,
                    AudioStatus = persisted.EndLesson.AudioStatus,
                    CoolingStatus = persisted.EndLesson.CoolingStatus,
                    DevicesPoweredOffStatus = persisted.EndLesson.DevicesPoweredOffStatus,
                    SeatingOrderStatus = persisted.EndLesson.SeatingOrderStatus,
                    RoomHygieneStatus = persisted.EndLesson.RoomHygieneStatus,
                    StudentRuleComplianceStatus = persisted.EndLesson.StudentRuleComplianceStatus,
                    ViolationListSummary = persisted.EndLesson.ViolationListSummary
                }
            };
        }

        private static ScheduleReportBundle MapReportsRequestToModel(
            ScheduleReportsRequest request,
            TeacherSchedule schedule,
            ScheduleRoomSessionContextResponse roomSessionContext)
        {
            request.StartLesson ??= new StartLessonReportRequest();
            request.Professional ??= new ProfessionalReportRequest();
            request.EndLesson ??= new EndLessonReportRequest();

            return new ScheduleReportBundle
            {
                StartLesson = new StartLessonReport
                {
                    TeacherName = Clean(request.StartLesson.TeacherName),
                    AssistantName = Clean(request.StartLesson.AssistantName),
                    RoomName = FirstNonEmpty(Clean(request.StartLesson.RoomName), schedule.RoomName),
                    TotalMachines = Clean(request.StartLesson.TotalMachines),
                    BrokenMachinesSummary = Clean(request.StartLesson.BrokenMachinesSummary),
                    MissingMachinesForStudents = Clean(request.StartLesson.MissingMachinesForStudents),
                    NetSupportStatus = Clean(request.StartLesson.NetSupportStatus),
                    AudioStatus = Clean(request.StartLesson.AudioStatus),
                    CoolingStatus = Clean(request.StartLesson.CoolingStatus),
                    HygieneStatus = Clean(request.StartLesson.HygieneStatus),
                },
                Professional = new ProfessionalReport
                {
                    TeacherName = Clean(request.Professional.TeacherName),
                    ClassName = FirstNonEmpty(Clean(request.Professional.ClassName), schedule.ClassName),
                    SubjectName = FirstNonEmpty(Clean(request.Professional.SubjectName), schedule.Subject),
                    TeachingMaterials = Clean(request.Professional.TeachingMaterials),
                    TeachingContent = Clean(request.Professional.TeachingContent),
                    PlannedLessons = Clean(request.Professional.PlannedLessons),
                    TaughtLessons = Clean(request.Professional.TaughtLessons),
                    OngoingPracticeCompletions = Clean(request.Professional.OngoingPracticeCompletions),
                    GmetrixResultRate = Clean(request.Professional.GmetrixResultRate),
                },
                EndLesson = new EndLessonReport
                {
                    TeacherName = Clean(request.EndLesson.TeacherName),
                    AssistantName = Clean(request.EndLesson.AssistantName),
                    RoomName = FirstNonEmpty(Clean(request.EndLesson.RoomName), schedule.RoomName),
                    TotalMachines = Clean(request.EndLesson.TotalMachines),
                    ClassStudentCountSummary = FirstNonEmpty(
                        Clean(request.EndLesson.ClassStudentCountSummary),
                        roomSessionContext.SharedClassStudentSummary),
                    StudentMaterialCoverageRate = Clean(request.EndLesson.StudentMaterialCoverageRate),
                    BrokenMachinesSummary = Clean(request.EndLesson.BrokenMachinesSummary),
                    NetSupportStatus = Clean(request.EndLesson.NetSupportStatus),
                    AudioStatus = Clean(request.EndLesson.AudioStatus),
                    CoolingStatus = Clean(request.EndLesson.CoolingStatus),
                    DevicesPoweredOffStatus = Clean(request.EndLesson.DevicesPoweredOffStatus),
                    SeatingOrderStatus = Clean(request.EndLesson.SeatingOrderStatus),
                    RoomHygieneStatus = Clean(request.EndLesson.RoomHygieneStatus),
                    StudentRuleComplianceStatus = Clean(request.EndLesson.StudentRuleComplianceStatus),
                    ViolationListSummary = Clean(request.EndLesson.ViolationListSummary),
                }
            };
        }

        private async Task<Class?> ResolveClassAsync(TeacherSchedule schedule, string ownerId, bool isAdmin)
        {
            if (!string.IsNullOrWhiteSpace(schedule.ClassId))
            {
                var byIdFilter = Builders<Class>.Filter.And(
                    Builders<Class>.Filter.Eq(c => c.Id, schedule.ClassId),
                    Builders<Class>.Filter.Eq(c => c.IsActive, true));

                if (!string.IsNullOrWhiteSpace(schedule.SchoolId))
                {
                    byIdFilter &= Builders<Class>.Filter.Eq(c => c.SchoolId, schedule.SchoolId);
                }

                return await _classes.Find(byIdFilter).FirstOrDefaultAsync();
            }

            if (string.IsNullOrWhiteSpace(schedule.ClassName))
            {
                return null;
            }

            var filter = Builders<Class>.Filter.And(
                Builders<Class>.Filter.Eq(x => x.Name, schedule.ClassName.Trim()),
                Builders<Class>.Filter.Eq(x => x.IsActive, true)
            );

            if (!isAdmin)
            {
                filter &= Builders<Class>.Filter.Eq(x => x.OwnerId, ownerId);
            }

            return await _classes
                .Find(filter)
                .SortByDescending(c => c.CreatedAt)
                .FirstOrDefaultAsync();
        }

        private static string NormalizeStatus(string? rawStatus)
        {
            if (string.IsNullOrWhiteSpace(rawStatus))
            {
                return AttendanceStatus.Present;
            }

            var normalized = rawStatus.Trim();
            if (string.Equals(normalized, AttendanceStatus.Present, StringComparison.OrdinalIgnoreCase))
            {
                return AttendanceStatus.Present;
            }

            if (string.Equals(normalized, AttendanceStatus.Absent, StringComparison.OrdinalIgnoreCase))
            {
                return AttendanceStatus.Absent;
            }

            return AttendanceStatus.Present;
        }

        private static int ParseTimeToMinutes(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return -1;
            }

            if (!TimeSpan.TryParse(value, out var time))
            {
                return -1;
            }

            return (int)time.TotalMinutes;
        }

        private static string ResolveSessionCode(string? startTime)
        {
            var minutes = ParseTimeToMinutes(startTime);
            if (minutes < 0)
            {
                return "Unknown";
            }

            if (minutes < 12 * 60)
            {
                return "Morning";
            }

            if (minutes < 18 * 60)
            {
                return "Afternoon";
            }

            return "Evening";
        }

        private static string ResolveSessionLabel(string? startTime)
        {
            return ResolveSessionCode(startTime) switch
            {
                "Morning" => "Buoi sang",
                "Afternoon" => "Buoi chieu",
                "Evening" => "Buoi toi",
                _ => "Khong xac dinh"
            };
        }

        private static string BuildClassStudentSummaryText(IEnumerable<ScheduleRoomClassSummaryResponse> classes)
        {
            var parts = classes.Select(x =>
            {
                if (x.MaxStudents.HasValue)
                {
                    return $"{x.ClassName}({x.CurrentStudents}/{x.MaxStudents.Value})";
                }

                return $"{x.ClassName}({x.CurrentStudents})";
            });

            return string.Join(" ", parts);
        }

        private static string NormalizeKey(string? value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : value.Trim().ToLowerInvariant();
        }

        private static string GetOwnerForSchedule(TeacherSchedule schedule, string fallbackOwnerId)
        {
            return string.IsNullOrWhiteSpace(schedule.OwnerId)
                ? fallbackOwnerId
                : schedule.OwnerId;
        }

        private static string Clean(string? value)
        {
            return value?.Trim() ?? string.Empty;
        }

        private static string FirstNonEmpty(string? preferred, string? fallback)
        {
            if (!string.IsNullOrWhiteSpace(preferred))
            {
                return preferred.Trim();
            }

            return fallback?.Trim() ?? string.Empty;
        }

        private static ScheduleAttendanceResponse BuildResponse(
            TeacherSchedule schedule,
            string classId,
            IReadOnlyCollection<ScheduleAttendanceStudentResponse> rows,
            ScheduleReportsResponse reports,
            ScheduleRoomSessionContextResponse roomSessionContext)
        {
            return new ScheduleAttendanceResponse
            {
                ScheduleId = schedule.Id ?? string.Empty,
                SchoolId = schedule.SchoolId,
                ClassId = classId,
                ClassName = schedule.ClassName,
                Subject = schedule.Subject,
                Date = schedule.Date,
                StartTime = schedule.StartTime,
                EndTime = schedule.EndTime,
                RoomName = schedule.RoomName,
                Students = rows.ToList(),
                PresentCount = rows.Count(x => x.AttendanceStatus == AttendanceStatus.Present),
                AbsentCount = rows.Count(x => x.AttendanceStatus == AttendanceStatus.Absent),
                Reports = reports,
                RoomSessionContext = roomSessionContext
            };
        }
    }
}
