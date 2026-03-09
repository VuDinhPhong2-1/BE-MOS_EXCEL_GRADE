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
        private readonly ILogger<AttendanceService> _logger;
        private static int _indexInitialized;

        public AttendanceService(IMongoDatabase database, ILogger<AttendanceService> logger)
        {
            _attendances = database.GetCollection<StudentScheduleAttendance>("scheduleAttendances");
            _students = database.GetCollection<Student>("students");
            _classes = database.GetCollection<Class>("Classes");
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
                throw new InvalidOperationException("Lịch dạy chưa gắn lớp hợp lệ, không thể điểm danh.");
            }
            if (string.IsNullOrWhiteSpace(classInfo.Id))
            {
                throw new InvalidOperationException("Không tìm thấy mã lớp hợp lệ để điểm danh.");
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

            return BuildResponse(schedule, classInfo.Id, rows);
        }

        public async Task<ScheduleAttendanceResponse> SaveScheduleAttendanceAsync(
            TeacherSchedule schedule,
            string ownerId,
            bool isAdmin,
            IReadOnlyCollection<SaveScheduleAttendanceItem> items,
            string updatedBy)
        {
            var classInfo = await ResolveClassAsync(schedule, ownerId, isAdmin);
            if (classInfo == null)
            {
                throw new InvalidOperationException("Lịch dạy chưa gắn lớp hợp lệ, không thể lưu điểm danh.");
            }
            if (string.IsNullOrWhiteSpace(classInfo.Id))
            {
                throw new InvalidOperationException("Không tìm thấy mã lớp hợp lệ để lưu điểm danh.");
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

            return await GetScheduleAttendanceAsync(schedule, ownerId, isAdmin);
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

        private static ScheduleAttendanceResponse BuildResponse(
            TeacherSchedule schedule,
            string classId,
            IReadOnlyCollection<ScheduleAttendanceStudentResponse> rows)
        {
            return new ScheduleAttendanceResponse
            {
                ScheduleId = schedule.Id ?? string.Empty,
                ClassId = classId,
                ClassName = schedule.ClassName,
                Subject = schedule.Subject,
                Date = schedule.Date,
                StartTime = schedule.StartTime,
                EndTime = schedule.EndTime,
                RoomName = schedule.RoomName,
                Students = rows.ToList(),
                PresentCount = rows.Count(x => x.AttendanceStatus == AttendanceStatus.Present),
                AbsentCount = rows.Count(x => x.AttendanceStatus == AttendanceStatus.Absent)
            };
        }
    }
}
