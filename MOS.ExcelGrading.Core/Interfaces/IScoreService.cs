// MOS.ExcelGrading.Core/Interfaces/IScoreService.cs
using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IScoreService
    {
        Task<List<ScoreResponse>> GetScoresByAssignmentAsync(string assignmentId);
        Task<List<ScoreResponse>> GetScoresByStudentAsync(string studentId);
        Task<StudentScoreReportResponse> GetStudentScoreReportAsync(string studentId, string classId);
        Task<Score?> GetScoreAsync(string studentId, string assignmentId);
        Task<Score> CreateOrUpdateScoreAsync(CreateScoreRequest request, string gradedBy);
        Task<List<Score>> BulkCreateOrUpdateScoresAsync(BulkScoreRequest request, string gradedBy);
        Task<bool> DeleteScoreAsync(string id, string userId);
    }
}
