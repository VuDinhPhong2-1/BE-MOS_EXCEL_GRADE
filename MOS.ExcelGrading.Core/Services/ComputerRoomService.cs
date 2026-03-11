using MongoDB.Driver;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public class ComputerRoomService : IComputerRoomService
    {
        private readonly IMongoCollection<ComputerRoom> _rooms;
        private static int _indexInitialized;

        public ComputerRoomService(IMongoDatabase database)
        {
            _rooms = database.GetCollection<ComputerRoom>("ComputerRooms");

            if (Interlocked.Exchange(ref _indexInitialized, 1) == 0)
            {
                var schoolNameIndex = new CreateIndexModel<ComputerRoom>(
                    Builders<ComputerRoom>.IndexKeys
                        .Ascending(x => x.SchoolId)
                        .Ascending(x => x.Name));

                var ownerSchoolIndex = new CreateIndexModel<ComputerRoom>(
                    Builders<ComputerRoom>.IndexKeys
                        .Ascending(x => x.OwnerId)
                        .Ascending(x => x.SchoolId)
                        .Ascending(x => x.IsActive));

                _rooms.Indexes.CreateMany(new[] { schoolNameIndex, ownerSchoolIndex });
            }
        }

        public async Task<List<ComputerRoom>> GetBySchoolIdAsync(string schoolId, bool includeInactive = false)
        {
            var filter = Builders<ComputerRoom>.Filter.Eq(x => x.SchoolId, schoolId);
            if (!includeInactive)
            {
                filter &= Builders<ComputerRoom>.Filter.Eq(x => x.IsActive, true);
            }

            return await _rooms
                .Find(filter)
                .SortBy(x => x.Name)
                .ThenBy(x => x.CreatedAt)
                .ToListAsync();
        }

        public async Task<ComputerRoom?> GetByIdAsync(string id)
        {
            return await _rooms.Find(x => x.Id == id).FirstOrDefaultAsync();
        }

        public async Task<ComputerRoom?> GetBySchoolAndNameAsync(string schoolId, string roomName, bool includeInactive = false)
        {
            if (string.IsNullOrWhiteSpace(schoolId) || string.IsNullOrWhiteSpace(roomName))
            {
                return null;
            }

            var normalized = roomName.Trim();
            var filter = Builders<ComputerRoom>.Filter.And(
                Builders<ComputerRoom>.Filter.Eq(x => x.SchoolId, schoolId),
                Builders<ComputerRoom>.Filter.Eq(x => x.Name, normalized));

            if (!includeInactive)
            {
                filter &= Builders<ComputerRoom>.Filter.Eq(x => x.IsActive, true);
            }

            return await _rooms.Find(filter).FirstOrDefaultAsync();
        }

        public async Task<ComputerRoom> CreateAsync(ComputerRoom room)
        {
            room.Name = room.Name.Trim();
            room.CreatedAt = DateTime.UtcNow;
            await _rooms.InsertOneAsync(room);
            return room;
        }

        public async Task<ComputerRoom?> UpdateAsync(string id, ComputerRoom room, string updatedBy, bool isAdmin)
        {
            var filter = isAdmin
                ? Builders<ComputerRoom>.Filter.Eq(x => x.Id, id)
                : Builders<ComputerRoom>.Filter.And(
                    Builders<ComputerRoom>.Filter.Eq(x => x.Id, id),
                    Builders<ComputerRoom>.Filter.Eq(x => x.OwnerId, updatedBy));

            var now = DateTime.UtcNow;
            var update = Builders<ComputerRoom>.Update
                .Set(x => x.Name, room.Name.Trim())
                .Set(x => x.StudentMachineCount, room.StudentMachineCount)
                .Set(x => x.TeacherMachineCount, room.TeacherMachineCount)
                .Set(x => x.BrokenMachineCount, room.BrokenMachineCount)
                .Set(x => x.NetSupportStatus, room.NetSupportStatus)
                .Set(x => x.AudioStatus, room.AudioStatus)
                .Set(x => x.CoolingStatus, room.CoolingStatus)
                .Set(x => x.DevicesPoweredOffStatus, room.DevicesPoweredOffStatus)
                .Set(x => x.SeatingOrderStatus, room.SeatingOrderStatus)
                .Set(x => x.RoomHygieneStatus, room.RoomHygieneStatus)
                .Set(x => x.IsActive, room.IsActive)
                .Set(x => x.UpdatedAt, now)
                .Set(x => x.UpdatedBy, updatedBy);

            return await _rooms.FindOneAndUpdateAsync(
                filter,
                update,
                new FindOneAndUpdateOptions<ComputerRoom> { ReturnDocument = ReturnDocument.After });
        }

        public async Task<bool> DeleteAsync(string id, string deletedBy, bool isAdmin)
        {
            var filter = isAdmin
                ? Builders<ComputerRoom>.Filter.Eq(x => x.Id, id)
                : Builders<ComputerRoom>.Filter.And(
                    Builders<ComputerRoom>.Filter.Eq(x => x.Id, id),
                    Builders<ComputerRoom>.Filter.Eq(x => x.OwnerId, deletedBy));

            var result = await _rooms.DeleteOneAsync(filter);
            return result.DeletedCount > 0;
        }
    }
}
