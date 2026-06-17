using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IExamSessionService
    {
        Task<ExamSession> StartSessionAsync(string publicationToken, StartExamSessionRequest request);
        Task<ExamSessionStateDto?> GetStateAsync(string publicationToken, string sessionId);
        Task<ExamSessionProjectBootstrapDto?> GetCurrentProjectBootstrapAsync(string sessionId);
        Task<ExamSessionProjectBootstrapDto?> GetCurrentProjectBootstrapAsync(string publicationToken, string sessionId);
        Task UploadScoreAsync(string sessionId, string projectCode, ScoreUploadRequest request);
        Task UploadScoreAsync(string publicationToken, string sessionId, string projectCode, ScoreUploadRequest request);
        Task<AdvanceExamSessionResponse> AdvanceAsync(string sessionId);
        Task<AdvanceExamSessionResponse> AdvanceAsync(string publicationToken, string sessionId);
        Task<RestartCurrentProjectResponse> RestartCurrentProjectAsync(string sessionId);
        Task<RestartCurrentProjectResponse> RestartCurrentProjectAsync(string publicationToken, string sessionId);
    }
}
