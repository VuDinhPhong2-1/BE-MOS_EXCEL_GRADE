using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T6Grader : ITaskGrader
    {
        public string TaskId => "P03-T6";
        public string TaskName => "Them Quick Print vao Quick Access Toolbar";
        public decimal MaxScore => 0;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore,
                Score = 0
            };

            result.Details.Add("Task nay khong the auto-grade tu file .xlsx.");
            result.Details.Add("Quick Access Toolbar la thiet lap UI cua ung dung Excel, khong luu trong workbook.");
            result.Details.Add("Can cham thu cong khi can xac nhan Task 6.");

            return result;
        }
    }
}

