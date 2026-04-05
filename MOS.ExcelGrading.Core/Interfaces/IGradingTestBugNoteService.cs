using MOS.ExcelGrading.Core.DTOs;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IGradingTestBugNoteService
    {
        Task<List<GradingTestBugNoteResponse>> GetByUserAsync(string userId, string? projectCode = null);
        Task<GradingTestBugNoteResponse> CreateAsync(CreateGradingTestBugNoteRequest request, string userId);
        Task<bool> DeleteAsync(string id, string userId);
    }
}
