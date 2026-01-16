using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IClassService
    {
        Task<Class> CreateClassAsync(Class classEntity, string ownerId);
        Task<Class?> GetClassByIdAsync(string id);
        Task<List<Class>> GetAllClassesAsync(bool includeInactive = false);
        Task<List<Class>> GetClassesBySchoolIdAsync(string schoolId, bool includeInactive = false);
        Task<List<Class>> GetClassesByOwnerIdAsync(string ownerId, bool includeInactive = false);
        Task<Class?> UpdateClassAsync(string id, Class classEntity, string updatedBy);
        Task<bool> DeleteClassAsync(string id);
        Task<bool> ClassExistsAsync( string schoolId);
        Task<bool> IsOwnerOfClassAsync(string userId, string classId);
        Task<bool> AddStudentToClassAsync(string classId, string studentId);
        Task<bool> RemoveStudentFromClassAsync(string classId, string studentId);
    }
}
