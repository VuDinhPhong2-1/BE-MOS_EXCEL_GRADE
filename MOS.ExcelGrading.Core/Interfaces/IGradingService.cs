using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IGradingService
    {
        Task<GradingResult> GradeProject01Async(Stream studentFile);
        Task<GradingResult> GradeProject02Async(Stream studentFile);
        Task<GradingResult> GradeProject03Async(Stream studentFile);
        Task<GradingResult> GradeProject04Async(Stream studentFile);
        Task<GradingResult> GradeProject09Async(Stream studentFile);
    }
}
