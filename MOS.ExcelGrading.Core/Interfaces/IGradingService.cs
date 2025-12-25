using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IGradingService
    {
        Task<GradingResult> GradeProject09Async(Stream studentFile, Stream answerFile);
    }
}
