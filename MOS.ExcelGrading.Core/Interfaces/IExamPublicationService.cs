using MOS.ExcelGrading.Core.DTOs;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IExamPublicationService
    {
        Task<ExamPublication> CreateExamPublicationAsync(CreateExamPublicationRequest request, string userId);
        Task<ExamPublication?> GetExamPublicationByIdAsync(string id);
        Task<PublicExamPublicationInfoDto?> GetPublicExamPublicationByTokenAsync(string publicationToken);
    }
}
