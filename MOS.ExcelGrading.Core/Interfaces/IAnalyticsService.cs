using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IAnalyticsService
    {
        Task SaveGradingAttemptAsync(
            GradingResult result,
            string projectEndpoint,
            string? classId,
            string? assignmentId,
            string? studentId,
            string? gradedBy,
            bool persistToDatabase = false);

        Task<ClassAnalyticsOverviewResponse> GetClassOverviewAsync(string classId);
        Task<List<WeakTaskResponse>> GetWeakTasksAsync(string classId, string? projectEndpoint, int top);
        Task<List<ProjectPerformanceResponse>> GetProjectPerformanceAsync(string classId);
    }
}
