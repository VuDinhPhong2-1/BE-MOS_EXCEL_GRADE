using Microsoft.Extensions.Logging;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class ScheduleService : IScheduleService
    {
        private readonly IMongoCollection<TeacherSchedule> _schedules;
        private static int _indexInitialized;
        private readonly ILogger<ScheduleService> _logger;

        public ScheduleService(IMongoDatabase database, ILogger<ScheduleService> logger)
        {
            _schedules = database.GetCollection<TeacherSchedule>("teacherSchedules");
            _logger = logger;

            if (Interlocked.Exchange(ref _indexInitialized, 1) == 0)
            {
                var ownerDateIndex = new CreateIndexModel<TeacherSchedule>(
                    Builders<TeacherSchedule>.IndexKeys
                        .Ascending(x => x.OwnerId)
                        .Ascending(x => x.Date)
                        .Ascending(x => x.StartTime));
                _schedules.Indexes.CreateOne(ownerDateIndex);
            }
        }

        public async Task<List<TeacherSchedule>> GetSchedulesByWeekAsync(string ownerId, DateTime weekStart, bool includeInactive = false)
        {
            var weekStartUtc = EnsureUtc(weekStart);
            var weekEndUtc = weekStartUtc.AddDays(7);

            var filter = Builders<TeacherSchedule>.Filter.And(
                Builders<TeacherSchedule>.Filter.Eq(x => x.OwnerId, ownerId),
                Builders<TeacherSchedule>.Filter.Gte(x => x.Date, weekStartUtc),
                Builders<TeacherSchedule>.Filter.Lt(x => x.Date, weekEndUtc)
            );

            if (!includeInactive)
            {
                filter &= Builders<TeacherSchedule>.Filter.Eq(x => x.IsActive, true);
            }

            return await _schedules
                .Find(filter)
                .SortBy(x => x.Date)
                .ThenBy(x => x.StartTime)
                .ToListAsync();
        }

        public async Task<TeacherSchedule?> GetByIdAsync(string id)
        {
            return await _schedules.Find(x => x.Id == id).FirstOrDefaultAsync();
        }

        public async Task<TeacherSchedule> CreateAsync(TeacherSchedule schedule)
        {
            schedule.CreatedAt = DateTime.UtcNow;
            schedule.Date = EnsureUtc(schedule.Date);
            await _schedules.InsertOneAsync(schedule);
            return schedule;
        }

        public async Task<TeacherSchedule?> UpdateAsync(string id, TeacherSchedule schedule, string updatedBy, bool isAdmin)
        {
            schedule.UpdatedAt = DateTime.UtcNow;
            schedule.UpdatedBy = updatedBy;
            schedule.Date = EnsureUtc(schedule.Date);

            var filter = isAdmin
                ? Builders<TeacherSchedule>.Filter.Eq(x => x.Id, id)
                : Builders<TeacherSchedule>.Filter.And(
                    Builders<TeacherSchedule>.Filter.Eq(x => x.Id, id),
                    Builders<TeacherSchedule>.Filter.Eq(x => x.OwnerId, updatedBy));

            var updates = new List<UpdateDefinition<TeacherSchedule>>
            {
                Builders<TeacherSchedule>.Update.Set(x => x.SchoolId, schedule.SchoolId),
                Builders<TeacherSchedule>.Update.Set(x => x.ClassId, schedule.ClassId),
                Builders<TeacherSchedule>.Update.Set(x => x.ClassName, schedule.ClassName),
                Builders<TeacherSchedule>.Update.Set(x => x.Subject, schedule.Subject),
                Builders<TeacherSchedule>.Update.Set(x => x.RoomName, schedule.RoomName),
                Builders<TeacherSchedule>.Update.Set(x => x.RoomId, schedule.RoomId),
                Builders<TeacherSchedule>.Update.Set(x => x.PeriodLabel, schedule.PeriodLabel),
                Builders<TeacherSchedule>.Update.Set(x => x.Date, schedule.Date),
                Builders<TeacherSchedule>.Update.Set(x => x.StartTime, schedule.StartTime),
                Builders<TeacherSchedule>.Update.Set(x => x.EndTime, schedule.EndTime),
                Builders<TeacherSchedule>.Update.Set(x => x.Notes, schedule.Notes),
                Builders<TeacherSchedule>.Update.Set(x => x.IsActive, schedule.IsActive),
                Builders<TeacherSchedule>.Update.Set(x => x.UpdatedAt, schedule.UpdatedAt),
                Builders<TeacherSchedule>.Update.Set(x => x.UpdatedBy, schedule.UpdatedBy)
            };

            if (!string.IsNullOrWhiteSpace(schedule.OwnerId))
            {
                updates.Add(Builders<TeacherSchedule>.Update.Set(x => x.OwnerId, schedule.OwnerId));
            }

            var update = Builders<TeacherSchedule>.Update.Combine(updates);

            var result = await _schedules.FindOneAndUpdateAsync(
                filter,
                update,
                new FindOneAndUpdateOptions<TeacherSchedule> { ReturnDocument = ReturnDocument.After });

            return result;
        }

        public async Task<bool> DeleteAsync(string id, string ownerId, bool isAdmin)
        {
            var filter = isAdmin
                ? Builders<TeacherSchedule>.Filter.Eq(x => x.Id, id)
                : Builders<TeacherSchedule>.Filter.And(
                    Builders<TeacherSchedule>.Filter.Eq(x => x.Id, id),
                    Builders<TeacherSchedule>.Filter.Eq(x => x.OwnerId, ownerId));

            var result = await _schedules.DeleteOneAsync(filter);
            if (result.DeletedCount == 0)
            {
                _logger.LogWarning("Không tìm thấy lịch để xóa: {ScheduleId}", id);
            }
            return result.DeletedCount > 0;
        }

        private static DateTime EnsureUtc(DateTime value)
        {
            return value.Kind switch
            {
                DateTimeKind.Utc => value,
                DateTimeKind.Local => value.ToUniversalTime(),
                _ => DateTime.SpecifyKind(value, DateTimeKind.Utc)
            };
        }
    }
}
