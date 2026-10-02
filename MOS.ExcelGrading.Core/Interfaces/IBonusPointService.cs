// MOS.ExcelGrading.Core/Interfaces/IBonusPointService.cs
using MOS.ExcelGrading.Core.DTOs;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IBonusPointService
    {
        Task<BonusPointResponse> CreateAsync(CreateBonusPointRequest request, string userId);
        Task<List<BonusPointResponse>> BulkCreateAsync(BulkBonusPointRequest request, string userId);
        Task<ClassBonusSummaryResponse> GetByClassAsync(string classId);
        Task<List<BonusPointResponse>> GetByClassAndDateAsync(string classId, DateTime date);
        Task<List<BonusPointResponse>> GetByStudentAsync(string studentId, string classId);
        Task<BonusPointResponse?> UpdateAsync(string id, UpdateBonusPointRequest request, string userId);
        Task<bool> DeleteAsync(string id, string userId);
    }
}
