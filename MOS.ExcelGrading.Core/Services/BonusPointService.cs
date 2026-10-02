// MOS.ExcelGrading.Core/Services/BonusPointService.cs
using Microsoft.Extensions.Logging;
using MongoDB.Bson;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class BonusPointService : IBonusPointService
    {
        private readonly IMongoCollection<BonusPoint> _bonusPoints;
        private readonly IMongoCollection<Student> _students;
        private readonly IMongoCollection<User> _users;
        private readonly ILogger<BonusPointService> _logger;
        private static int _indexInitialized;

        public BonusPointService(
            IMongoDatabase database,
            ILogger<BonusPointService> logger)
        {
            _bonusPoints = database.GetCollection<BonusPoint>("bonus_points");
            _students = database.GetCollection<Student>("students");
            _users = database.GetCollection<User>("users");
            _logger = logger;

            if (Interlocked.Exchange(ref _indexInitialized, 1) == 0)
            {
                try
                {
                    var classDateIndex = new CreateIndexModel<BonusPoint>(
                        Builders<BonusPoint>.IndexKeys
                            .Ascending(x => x.ClassId)
                            .Descending(x => x.Date));

                    var classStudentIndex = new CreateIndexModel<BonusPoint>(
                        Builders<BonusPoint>.IndexKeys
                            .Ascending(x => x.ClassId)
                            .Ascending(x => x.StudentId));

                    _bonusPoints.Indexes.CreateMany(new[] { classDateIndex, classStudentIndex });
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "⚠️ Could not initialize indexes for bonus_points collection");
                }
            }
        }

        public async Task<BonusPointResponse> CreateAsync(CreateBonusPointRequest request, string userId)
        {
            EnsureValidObjectId(request.ClassId, "Mã lớp");
            EnsureValidObjectId(request.StudentId, "Mã học sinh");

            var category = string.IsNullOrWhiteSpace(request.Category) ? BonusPointCategories.Participation : request.Category.Trim();

            var entity = new BonusPoint
            {
                ClassId = request.ClassId,
                StudentId = request.StudentId,
                Date = request.Date.ToUniversalTime(),
                Points = request.Points,
                Category = category,
                Reason = request.Reason?.Trim(),
                ScheduleId = string.IsNullOrWhiteSpace(request.ScheduleId) ? null : request.ScheduleId,
                CreatedAt = DateTime.UtcNow,
                CreatedBy = userId
            };

            await _bonusPoints.InsertOneAsync(entity);

            var student = await SafeGetStudentAsync(entity.StudentId);
            var creatorName = await ResolveUserNameAsync(userId);

            return BuildBonusPointResponse(entity, student, creatorName);
        }

        public async Task<List<BonusPointResponse>> BulkCreateAsync(BulkBonusPointRequest request, string userId)
        {
            EnsureValidObjectId(request.ClassId, "Mã lớp");

            if (request.Items == null || request.Items.Count == 0)
            {
                return new List<BonusPointResponse>();
            }

            var entities = new List<BonusPoint>();
            var utcDate = request.Date.ToUniversalTime();

            foreach (var item in request.Items)
            {
                if (string.IsNullOrWhiteSpace(item.StudentId))
                    continue;

                var category = string.IsNullOrWhiteSpace(item.Category) ? BonusPointCategories.Participation : item.Category.Trim();

                entities.Add(new BonusPoint
                {
                    ClassId = request.ClassId,
                    StudentId = item.StudentId,
                    Date = utcDate,
                    Points = item.Points,
                    Category = category,
                    Reason = item.Reason?.Trim(),
                    ScheduleId = string.IsNullOrWhiteSpace(request.ScheduleId) ? null : request.ScheduleId,
                    CreatedAt = DateTime.UtcNow,
                    CreatedBy = userId
                });
            }

            if (entities.Count > 0)
            {
                await _bonusPoints.InsertManyAsync(entities);
            }

            var students = await SafeGetStudentsByClassAsync(request.ClassId);
            var studentsById = students.ToDictionary(s => s.Id ?? string.Empty, StringComparer.Ordinal);
            var creatorName = await ResolveUserNameAsync(userId);

            return entities.Select(e =>
            {
                studentsById.TryGetValue(e.StudentId, out var student);
                return BuildBonusPointResponse(e, student, creatorName);
            }).ToList();
        }

        public async Task<ClassBonusSummaryResponse> GetByClassAsync(string classId)
        {
            EnsureValidObjectId(classId, "Mã lớp");

            var students = await SafeGetStudentsByClassAsync(classId);
            var allPoints = await _bonusPoints
                .Find(b => b.ClassId == classId)
                .SortByDescending(b => b.Date)
                .ToListAsync();

            var pointsByStudent = allPoints
                .GroupBy(p => p.StudentId)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);

            // Pre-cache creator names
            var creatorIds = allPoints
                .Where(p => !string.IsNullOrWhiteSpace(p.CreatedBy))
                .Select(p => p.CreatedBy!)
                .Distinct()
                .ToList();
            var creatorNames = await ResolveUserNamesAsync(creatorIds);

            var studentSummaries = new List<StudentBonusPointSummary>();

            foreach (var student in students)
            {
                var sId = student.Id ?? string.Empty;
                pointsByStudent.TryGetValue(sId, out var studentPoints);
                studentPoints ??= new List<BonusPoint>();

                var details = studentPoints.Select(p =>
                {
                    creatorNames.TryGetValue(p.CreatedBy ?? string.Empty, out var cName);
                    return BuildBonusPointResponse(p, student, cName);
                }).ToList();

                var total = studentPoints.Sum(p => p.Points);

                studentSummaries.Add(new StudentBonusPointSummary
                {
                    StudentId = sId,
                    StudentFirstName = student.FirstName ?? string.Empty,
                    StudentMiddleName = student.MiddleName ?? string.Empty,
                    StudentFullName = $"{student.MiddleName} {student.FirstName}".Trim(),
                    TotalBonusPoints = total,
                    Details = details
                });
            }

            // Also check for any bonus points belonging to students not in current class list
            foreach (var kvp in pointsByStudent)
            {
                if (students.Any(s => s.Id == kvp.Key))
                    continue;

                var student = await SafeGetStudentAsync(kvp.Key);
                var details = kvp.Value.Select(p =>
                {
                    creatorNames.TryGetValue(p.CreatedBy ?? string.Empty, out var cName);
                    return BuildBonusPointResponse(p, student, cName);
                }).ToList();

                studentSummaries.Add(new StudentBonusPointSummary
                {
                    StudentId = kvp.Key,
                    StudentFirstName = student?.FirstName ?? string.Empty,
                    StudentMiddleName = student?.MiddleName ?? string.Empty,
                    StudentFullName = $"{student?.MiddleName} {student?.FirstName}".Trim(),
                    TotalBonusPoints = kvp.Value.Sum(p => p.Points),
                    Details = details
                });
            }

            return new ClassBonusSummaryResponse
            {
                ClassId = classId,
                Students = studentSummaries.OrderBy(s => s.StudentFullName).ToList()
            };
        }

        public async Task<List<BonusPointResponse>> GetByClassAndDateAsync(string classId, DateTime date)
        {
            EnsureValidObjectId(classId, "Mã lớp");

            var startUtc = date.Date.ToUniversalTime();
            var endUtc = startUtc.AddDays(1);

            var points = await _bonusPoints
                .Find(b => b.ClassId == classId && b.Date >= startUtc && b.Date < endUtc)
                .SortBy(b => b.CreatedAt)
                .ToListAsync();

            var students = await SafeGetStudentsByClassAsync(classId);
            var studentsById = students.ToDictionary(s => s.Id ?? string.Empty, StringComparer.Ordinal);

            var creatorIds = points
                .Where(p => !string.IsNullOrWhiteSpace(p.CreatedBy))
                .Select(p => p.CreatedBy!)
                .Distinct()
                .ToList();
            var creatorNames = await ResolveUserNamesAsync(creatorIds);

            return points.Select(p =>
            {
                studentsById.TryGetValue(p.StudentId, out var student);
                creatorNames.TryGetValue(p.CreatedBy ?? string.Empty, out var cName);
                return BuildBonusPointResponse(p, student, cName);
            }).ToList();
        }

        public async Task<List<BonusPointResponse>> GetByStudentAsync(string studentId, string classId)
        {
            EnsureValidObjectId(studentId, "Mã học sinh");
            EnsureValidObjectId(classId, "Mã lớp");

            var points = await _bonusPoints
                .Find(b => b.StudentId == studentId && b.ClassId == classId)
                .SortByDescending(b => b.Date)
                .ToListAsync();

            var student = await SafeGetStudentAsync(studentId);

            var creatorIds = points
                .Where(p => !string.IsNullOrWhiteSpace(p.CreatedBy))
                .Select(p => p.CreatedBy!)
                .Distinct()
                .ToList();
            var creatorNames = await ResolveUserNamesAsync(creatorIds);

            return points.Select(p =>
            {
                creatorNames.TryGetValue(p.CreatedBy ?? string.Empty, out var cName);
                return BuildBonusPointResponse(p, student, cName);
            }).ToList();
        }

        public async Task<BonusPointResponse?> UpdateAsync(string id, UpdateBonusPointRequest request, string userId)
        {
            EnsureValidObjectId(id, "Mã điểm cộng");

            var updateBuilder = Builders<BonusPoint>.Update;
            var updates = new List<UpdateDefinition<BonusPoint>>();

            if (request.Date.HasValue)
                updates.Add(updateBuilder.Set(b => b.Date, request.Date.Value.ToUniversalTime()));

            if (request.Points.HasValue)
                updates.Add(updateBuilder.Set(b => b.Points, request.Points.Value));

            if (!string.IsNullOrWhiteSpace(request.Category))
                updates.Add(updateBuilder.Set(b => b.Category, request.Category.Trim()));

            if (request.Reason != null)
                updates.Add(updateBuilder.Set(b => b.Reason, request.Reason.Trim()));

            if (request.ScheduleId != null)
                updates.Add(updateBuilder.Set(b => b.ScheduleId, string.IsNullOrWhiteSpace(request.ScheduleId) ? null : request.ScheduleId));

            updates.Add(updateBuilder.Set(b => b.UpdatedAt, DateTime.UtcNow));
            updates.Add(updateBuilder.Set(b => b.UpdatedBy, userId));

            var result = await _bonusPoints.FindOneAndUpdateAsync(
                b => b.Id == id,
                updateBuilder.Combine(updates),
                new FindOneAndUpdateOptions<BonusPoint> { ReturnDocument = ReturnDocument.After }
            );

            if (result == null)
                return null;

            var student = await SafeGetStudentAsync(result.StudentId);
            var creatorName = await ResolveUserNameAsync(result.CreatedBy);

            return BuildBonusPointResponse(result, student, creatorName);
        }

        public async Task<bool> DeleteAsync(string id, string userId)
        {
            EnsureValidObjectId(id, "Mã điểm cộng");

            var result = await _bonusPoints.DeleteOneAsync(b => b.Id == id);
            return result.DeletedCount > 0;
        }

        // ========== HELPER METHODS ==========

        private static BonusPointResponse BuildBonusPointResponse(
            BonusPoint point,
            Student? student,
            string? creatorName)
        {
            return new BonusPointResponse
            {
                Id = point.Id,
                StudentId = point.StudentId,
                StudentFirstName = student?.FirstName ?? string.Empty,
                StudentMiddleName = student?.MiddleName ?? string.Empty,
                StudentFullName = $"{student?.MiddleName} {student?.FirstName}".Trim(),
                ClassId = point.ClassId,
                Date = point.Date,
                Points = point.Points,
                Category = point.Category,
                Reason = point.Reason,
                ScheduleId = point.ScheduleId,
                CreatedAt = point.CreatedAt,
                CreatedBy = point.CreatedBy,
                CreatedByName = creatorName,
                UpdatedAt = point.UpdatedAt
            };
        }

        private async Task<Student?> SafeGetStudentAsync(string studentId)
        {
            if (string.IsNullOrWhiteSpace(studentId) || !ObjectId.TryParse(studentId.Trim(), out _))
                return null;

            try
            {
                return await _students.Find(s => s.Id == studentId).FirstOrDefaultAsync();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "⚠️ Error fetching student {StudentId}", studentId);
                return null;
            }
        }

        private async Task<List<Student>> SafeGetStudentsByClassAsync(string classId)
        {
            try
            {
                return await _students.Find(s => s.ClassId == classId).ToListAsync();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "⚠️ Error fetching students for class {ClassId}", classId);
                return new List<Student>();
            }
        }

        private async Task<string?> ResolveUserNameAsync(string? userId)
        {
            if (string.IsNullOrWhiteSpace(userId) || !ObjectId.TryParse(userId.Trim(), out _))
                return null;

            try
            {
                var user = await _users.Find(u => u.Id == userId).FirstOrDefaultAsync();
                return user?.FullName ?? user?.Username;
            }
            catch
            {
                return null;
            }
        }

        private async Task<Dictionary<string, string>> ResolveUserNamesAsync(List<string> userIds)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            if (userIds == null || userIds.Count == 0)
                return result;

            var validObjectIds = userIds
                .Where(id => ObjectId.TryParse(id, out _))
                .Distinct()
                .ToList();

            if (validObjectIds.Count == 0)
                return result;

            try
            {
                var users = await _users.Find(u => validObjectIds.Contains(u.Id ?? string.Empty)).ToListAsync();
                foreach (var u in users)
                {
                    if (!string.IsNullOrWhiteSpace(u.Id))
                    {
                        result[u.Id] = u.FullName ?? u.Username;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "⚠️ Error resolving user names");
            }

            return result;
        }

        private static void EnsureValidObjectId(string value, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value) || !ObjectId.TryParse(value.Trim(), out _))
            {
                throw new ArgumentException($"{fieldName} không hợp lệ (cần ObjectId 24 ký tự).");
            }
        }
    }
}
