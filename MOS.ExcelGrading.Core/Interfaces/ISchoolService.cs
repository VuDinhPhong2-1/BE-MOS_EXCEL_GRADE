using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface ISchoolService
    {
        Task<School> CreateSchoolAsync(School school, string ownerId);
        Task<School?> GetSchoolByIdAsync(string id);
        Task<List<School>> GetAllSchoolsAsync(bool includeInactive = false);
        Task<List<School>> GetSchoolsByOwnerIdAsync(string ownerId, bool includeInactive = false);
        Task<School?> UpdateSchoolAsync(string id, School school, string updatedBy);
        Task<bool> DeleteSchoolAsync(string id);
        Task<bool> SchoolExistsAsync(string code);
        Task<bool> IsOwnerOfSchoolAsync(string userId, string schoolId);
    }
}
