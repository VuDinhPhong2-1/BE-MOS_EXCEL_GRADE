using Microsoft.AspNetCore.Http;
using MOS.ExcelGrading.Core.DTOs;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface ISubmissionPortalService
    {
        Task<SubmissionPortalResponse> CreateAsync(CreateSubmissionPortalRequest request, string userId);
        Task<List<SubmissionPortalResponse>> GetAllAsync(string userId, bool isAdmin);
        Task<SubmissionPortalResponse?> GetByIdAsync(string id);
        Task<SubmissionPortalResponse?> UpdateAsync(string id, UpdateSubmissionPortalRequest request, string userId);
        Task<bool> DeleteAsync(string id, string userId);
        Task<PublicPortalInfoResponse?> GetPublicInfoAsync(string token);
        Task<List<PublicPortalStudentDto>> GetPublicStudentsAsync(string token, string classId);
        Task<PublicPortalSubmitResult> GradeAndSubmitAsync(string token, string classId, string studentId, string assignmentId, IFormFile file, string? ipAddress, string? userAgent);
        Task<List<SubmissionLeaderboardItem>> GetLeaderboardAsync(string token, string? classId = null, string? assignmentId = null);
        Task<List<SubmissionAlertResponse>> GetAlertsAsync(string portalId, bool includeDismissed = false);
        Task<int> GetUnreadAlertCountAsync(string portalId);
        Task<bool> UpdateAlertAsync(string portalId, string alertId, bool? isRead, bool? isDismissed);
        Task<List<SubmissionLogResponse>> GetSubmissionLogsAsync(string portalId);
    }
}