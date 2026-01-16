using Microsoft.Extensions.Options;
using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class SchoolService : ISchoolService
    {
        private readonly IMongoCollection<School> _schools;

        public SchoolService(IOptions<MongoDbSettings> mongoSettings)
        {
            var client = new MongoClient(mongoSettings.Value.ConnectionString);
            var database = client.GetDatabase(mongoSettings.Value.DatabaseName);
            _schools = database.GetCollection<School>(mongoSettings.Value.SchoolsCollectionName);

            // Tạo index unique cho Code
            var indexKeysDefinition = Builders<School>.IndexKeys.Ascending(s => s.Code);
            var indexOptions = new CreateIndexOptions { Unique = true };
            var indexModel = new CreateIndexModel<School>(indexKeysDefinition, indexOptions);
            _schools.Indexes.CreateOne(indexModel);
        }

        public async Task<School> CreateSchoolAsync(School school, string ownerId)
        {
            school.OwnerId = ownerId;
            school.CreatedBy = ownerId;
            school.CreatedAt = DateTime.UtcNow;
            school.IsActive = true;

            await _schools.InsertOneAsync(school);
            return school;
        }

        public async Task<School?> GetSchoolByIdAsync(string id)
        {
            return await _schools.Find(s => s.Id == id).FirstOrDefaultAsync();
        }

        public async Task<List<School>> GetAllSchoolsAsync(bool includeInactive = false)
        {
            var filter = includeInactive
                ? Builders<School>.Filter.Empty
                : Builders<School>.Filter.Eq(s => s.IsActive, true);

            return await _schools.Find(filter).ToListAsync();
        }

        public async Task<List<School>> GetSchoolsByOwnerIdAsync(string ownerId, bool includeInactive = false)
        {
            var filterBuilder = Builders<School>.Filter;
            var filter = filterBuilder.Eq(s => s.OwnerId, ownerId);

            if (!includeInactive)
            {
                filter &= filterBuilder.Eq(s => s.IsActive, true);
            }

            return await _schools.Find(filter).ToListAsync();
        }

        public async Task<School?> UpdateSchoolAsync(string id, School school, string updatedBy)
        {
            school.UpdatedAt = DateTime.UtcNow;
            school.UpdatedBy = updatedBy;

            var update = Builders<School>.Update
                .Set(s => s.Name, school.Name)
                .Set(s => s.Address, school.Address)
                .Set(s => s.PhoneNumber, school.PhoneNumber)
                .Set(s => s.Email, school.Email)
                .Set(s => s.Website, school.Website)
                .Set(s => s.Description, school.Description)
                .Set(s => s.Logo, school.Logo)
                .Set(s => s.IsActive, school.IsActive)
                .Set(s => s.UpdatedAt, school.UpdatedAt)
                .Set(s => s.UpdatedBy, school.UpdatedBy);

            var result = await _schools.FindOneAndUpdateAsync(
                s => s.Id == id,
                update,
                new FindOneAndUpdateOptions<School> { ReturnDocument = ReturnDocument.After }
            );

            return result;
        }

        public async Task<bool> DeleteSchoolAsync(string id)
        {
            var result = await _schools.DeleteOneAsync(s => s.Id == id);
            return result.DeletedCount > 0;
        }

        public async Task<bool> SchoolExistsAsync(string code)
        {
            var count = await _schools.CountDocumentsAsync(s => s.Code == code);
            return count > 0;
        }

        public async Task<bool> IsOwnerOfSchoolAsync(string userId, string schoolId)
        {
            var school = await _schools.Find(s => s.Id == schoolId && s.OwnerId == userId).FirstOrDefaultAsync();
            return school != null;
        }
    }
}
