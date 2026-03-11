using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IComputerRoomService
    {
        Task<List<ComputerRoom>> GetBySchoolIdAsync(string schoolId, bool includeInactive = false);
        Task<ComputerRoom?> GetByIdAsync(string id);
        Task<ComputerRoom?> GetBySchoolAndNameAsync(string schoolId, string roomName, bool includeInactive = false);
        Task<ComputerRoom> CreateAsync(ComputerRoom room);
        Task<ComputerRoom?> UpdateAsync(string id, ComputerRoom room, string updatedBy, bool isAdmin);
        Task<bool> DeleteAsync(string id, string deletedBy, bool isAdmin);
    }
}
