using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface IWordTaskGrader
    {
        string TaskId { get; }
        string TaskName { get; }
        decimal MaxScore { get; }
        TaskResult Grade(WordGradingContext studentDocument, WordGradingContext? answerDocument = null);
    }
}
