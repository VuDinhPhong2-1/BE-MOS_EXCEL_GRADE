using OfficeOpenXml;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Interfaces
{
    public interface ITaskGrader
    {
        string TaskId { get; }
        string TaskName { get; }
        decimal MaxScore { get; }
        TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet);
    }
}
