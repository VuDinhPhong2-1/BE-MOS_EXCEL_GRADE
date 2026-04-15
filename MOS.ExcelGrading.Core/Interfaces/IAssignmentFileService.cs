using Microsoft.AspNetCore.Http;
using MOS.ExcelGrading.Core.DTOs;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IAssignmentFileService
    {
        Task<AssignmentFileResponse> UploadAssignmentFileAsync(
            string assignmentId,
            IFormFile file,
            string subject,
            string kind,
            string userId);

        Task<List<AssignmentFileResponse>> GetAssignmentFilesAsync(
            string assignmentId,
            bool includeInactive = false);

        Task<AssignmentFileResponse?> GetAssignmentFileByIdAsync(string fileId);

        Task<AssignmentFileDownloadResult> OpenAssignmentFileDownloadAsync(string fileId);

        Task<bool> SoftDeleteAssignmentFileAsync(string fileId, string userId);
    }
}
