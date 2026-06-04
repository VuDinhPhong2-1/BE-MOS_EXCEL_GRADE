using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project03
{
    public class P03T6Grader : ITaskGrader
    {
        public string TaskId => "P03-T6";
        public string TaskName => "Thêm Quick Print vào Quick Access Toolbar";
        public decimal MaxScore => 0;

        public TaskResult Grade(ExcelWorksheet studentSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore,
                Score = 0
            };

            result.Details.Add("Task nay không thể auto-grade từ file .xlsx.");
            result.Details.Add("Quick Access Toolbar là thiết lập UI của ứng dụng Excel, không lưu trong workbook.");
            result.Details.Add("ần chấm thủ công khi cần xác nhận Task 6.");

            return result;
        }
    }
}


// minor-sync: non-functional graders update
