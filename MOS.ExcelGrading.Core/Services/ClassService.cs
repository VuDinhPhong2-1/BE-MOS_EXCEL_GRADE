using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class ClassService : IClassService
    {
        private readonly IMongoCollection<Class> _classes;
        private readonly ILogger<ClassService> _logger;

        public ClassService(
            IOptions<MongoDbSettings> mongoSettings,
            ILogger<ClassService> logger)
        {
            _logger = logger;

            var client = new MongoClient(mongoSettings.Value.ConnectionString);
            var database = client.GetDatabase(mongoSettings.Value.DatabaseName);
            _classes = database.GetCollection<Class>(mongoSettings.Value.ClassesCollectionName);

            _logger.LogInformation("✅ ClassService initialized successfully");
        }

        public async Task<Class> CreateClassAsync(Class classEntity, string ownerId)
        {
            try
            {
                _logger.LogInformation($"📤 CreateClassAsync called: Name={classEntity.Name}, SchoolId={classEntity.SchoolId}, OwnerId={ownerId}");

                // ✅ KIỂM TRA TRÙNG TÊN TRƯỚC KHI TẠO
                var existingClass = await _classes
                    .Find(c => c.SchoolId == classEntity.SchoolId
                            && c.Name == classEntity.Name
                            && c.IsActive)
                    .FirstOrDefaultAsync();

                if (existingClass != null)
                {
                    _logger.LogWarning($"❌ Class already exists: Name={classEntity.Name}, SchoolId={classEntity.SchoolId}");
                    throw new InvalidOperationException($"Lớp '{classEntity.Name}' đã tồn tại trong trường này!");
                }

                classEntity.OwnerId = ownerId;
                classEntity.CreatedBy = ownerId;
                classEntity.CreatedAt = DateTime.UtcNow;
                classEntity.IsActive = true;

                await _classes.InsertOneAsync(classEntity);

                _logger.LogInformation($"✅ Class created successfully: Id={classEntity.Id}, Name={classEntity.Name}");

                return classEntity;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in CreateClassAsync: Name={classEntity.Name}, SchoolId={classEntity.SchoolId}");
                throw;
            }
        }

        public async Task<Class?> GetClassByIdAsync(string id)
        {
            try
            {
                _logger.LogInformation($"📤 GetClassByIdAsync called: Id={id}");

                // ✅ KIỂM TRA ID HỢP LỆ
                if (string.IsNullOrEmpty(id) || id.Length != 24)
                {
                    _logger.LogWarning($"❌ Invalid class ID format: {id}");
                    return null;
                }

                // ✅ TÌM CLASS TRONG DATABASE
                Class? classEntity = null;

                try
                {
                    classEntity = await _classes.Find(c => c.Id == id).FirstOrDefaultAsync();
                }
                catch (MongoDB.Bson.BsonSerializationException ex)
                {
                    _logger.LogError(ex, $"❌ Deserialization error for class ID: {id}. Database may contain fields not in model.");
                    throw new InvalidOperationException($"Lỗi đọc dữ liệu lớp học. Vui lòng liên hệ quản trị viên.", ex);
                }

                if (classEntity == null)
                {
                    _logger.LogWarning($"❌ Class not found: Id={id}");
                }
                else
                {
                    _logger.LogInformation($"✅ Class found: Id={id}, Name={classEntity.Name}");
                }

                return classEntity;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in GetClassByIdAsync: Id={id}");
                throw;
            }
        }

        public async Task<List<Class>> GetAllClassesAsync(bool includeInactive = false)
        {
            try
            {
                _logger.LogInformation($"📤 GetAllClassesAsync called: includeInactive={includeInactive}");

                var filter = includeInactive
                    ? Builders<Class>.Filter.Empty
                    : Builders<Class>.Filter.Eq(c => c.IsActive, true);

                var classes = await _classes.Find(filter).ToListAsync();

                _logger.LogInformation($"✅ Found {classes.Count} classes");

                return classes;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "❌ Error in GetAllClassesAsync");
                throw;
            }
        }

        public async Task<List<Class>> GetClassesBySchoolIdAsync(string schoolId, bool includeInactive = false)
        {
            try
            {
                _logger.LogInformation($"📤 GetClassesBySchoolIdAsync called: schoolId={schoolId}, includeInactive={includeInactive}");

                var filterBuilder = Builders<Class>.Filter;
                var filter = filterBuilder.Eq(c => c.SchoolId, schoolId);

                if (!includeInactive)
                {
                    filter &= filterBuilder.Eq(c => c.IsActive, true);
                }

                var classes = await _classes.Find(filter).ToListAsync();

                _logger.LogInformation($"✅ Found {classes.Count} classes for school {schoolId}");

                return classes;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in GetClassesBySchoolIdAsync: schoolId={schoolId}");
                throw;
            }
        }

        public async Task<List<Class>> GetClassesByOwnerIdAsync(string ownerId, bool includeInactive = false)
        {
            try
            {
                _logger.LogInformation($"📤 GetClassesByOwnerIdAsync called: ownerId={ownerId}, includeInactive={includeInactive}");

                var filterBuilder = Builders<Class>.Filter;
                var filter = filterBuilder.Eq(c => c.OwnerId, ownerId);

                if (!includeInactive)
                {
                    filter &= filterBuilder.Eq(c => c.IsActive, true);
                }

                var classes = await _classes.Find(filter).ToListAsync();

                _logger.LogInformation($"✅ Found {classes.Count} classes for owner {ownerId}");

                return classes;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in GetClassesByOwnerIdAsync: ownerId={ownerId}");
                throw;
            }
        }

        public async Task<bool> DeleteClassAsync(string id)
        {
            try
            {
                _logger.LogInformation($"📤 DeleteClassAsync called: Id={id}");

                // ✅ HARD DELETE (xóa hẳn khỏi database)
                var result = await _classes.DeleteOneAsync(c => c.Id == id);

                if (result.DeletedCount > 0)
                {
                    _logger.LogInformation($"✅ Class deleted successfully: Id={id}");
                }
                else
                {
                    _logger.LogWarning($"❌ Class not found: Id={id}");
                }

                return result.DeletedCount > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in DeleteClassAsync: Id={id}");
                throw;
            }
        }


        public async Task<bool> ClassExistsAsync(string Name)
        {
            try
            {

                var count = await _classes.CountDocumentsAsync(
                    c => c.Name == Name && c.IsActive
                );

                var exists = count > 0;

                _logger.LogInformation($"✅ ClassExistsAsync result: {exists} (count={count})");

                return exists;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in ClassExistsAsync: schoolId={Name}");
                throw;
            }
        }

        public async Task<bool> IsOwnerOfClassAsync(string userId, string classId)
        {
            try
            {
                _logger.LogInformation($"📤 IsOwnerOfClassAsync called: userId={userId}, classId={classId}");

                var classEntity = await _classes
                    .Find(c => c.Id == classId && c.OwnerId == userId)
                    .FirstOrDefaultAsync();

                var isOwner = classEntity != null;

                _logger.LogInformation($"✅ IsOwnerOfClassAsync result: {isOwner}");

                return isOwner;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in IsOwnerOfClassAsync: userId={userId}, classId={classId}");
                throw;
            }
        }

        public async Task<bool> AddStudentToClassAsync(string classId, string studentId)
        {
            try
            {
                _logger.LogInformation($"📤 AddStudentToClassAsync called: classId={classId}, studentId={studentId}");

                var update = Builders<Class>.Update
                    .AddToSet(c => c.StudentIds, studentId)
                    .Inc(c => c.CurrentStudents, 1);

                var result = await _classes.UpdateOneAsync(c => c.Id == classId, update);

                if (result.ModifiedCount > 0)
                {
                    _logger.LogInformation($"✅ Student added to class successfully: classId={classId}, studentId={studentId}");
                }
                else
                {
                    _logger.LogWarning($"❌ Failed to add student to class: classId={classId}, studentId={studentId}");
                }

                return result.ModifiedCount > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in AddStudentToClassAsync: classId={classId}, studentId={studentId}");
                throw;
            }
        }

        public async Task<bool> RemoveStudentFromClassAsync(string classId, string studentId)
        {
            try
            {
                _logger.LogInformation($"📤 RemoveStudentFromClassAsync called: classId={classId}, studentId={studentId}");

                var update = Builders<Class>.Update
                    .Pull(c => c.StudentIds, studentId)
                    .Inc(c => c.CurrentStudents, -1);

                var result = await _classes.UpdateOneAsync(c => c.Id == classId, update);

                if (result.ModifiedCount > 0)
                {
                    _logger.LogInformation($"✅ Student removed from class successfully: classId={classId}, studentId={studentId}");
                }
                else
                {
                    _logger.LogWarning($"❌ Failed to remove student from class: classId={classId}, studentId={studentId}");
                }

                return result.ModifiedCount > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in RemoveStudentFromClassAsync: classId={classId}, studentId={studentId}");
                throw;
            }
        }

        public async Task<Class?> UpdateClassAsync(string id, Class classEntity, string updatedBy)
        {
            try
            {
                _logger.LogInformation($"📤 UpdateClassAsync called: Id={id}, Name={classEntity.Name}");

                classEntity.UpdatedAt = DateTime.UtcNow;
                classEntity.UpdatedBy = updatedBy;

                var update = Builders<Class>.Update
                    .Set(c => c.Name, classEntity.Name)
                    .Set(c => c.Description, classEntity.Description)
                    .Set(c => c.MaxStudents, classEntity.MaxStudents)
                    .Set(c => c.AcademicYear, classEntity.AcademicYear)
                    .Set(c => c.Grade, classEntity.Grade)
                    .Set(c => c.IsActive, classEntity.IsActive)
                    .Set(c => c.UpdatedAt, classEntity.UpdatedAt)
                    .Set(c => c.UpdatedBy, classEntity.UpdatedBy);

                var result = await _classes.FindOneAndUpdateAsync(
                    c => c.Id == id,
                    update,
                    new FindOneAndUpdateOptions<Class> { ReturnDocument = ReturnDocument.After }
                );

                if (result != null)
                {
                    _logger.LogInformation($"✅ Class updated successfully: Id={id}, Name={result.Name}");
                }
                else
                {
                    _logger.LogWarning($"❌ Class not found for update: Id={id}");
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"❌ Error in UpdateClassAsync: Id={id}");
                throw;
            }
        }

    }
}
