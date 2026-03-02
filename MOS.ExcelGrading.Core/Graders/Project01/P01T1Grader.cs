using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project01
{
    public class P01T1Grader : ITaskGrader
    {
        public string TaskId => "P01-T1";
        public string TaskName => "Sao chép định dạng A1:A2 từ Documentation sang Menu Items";
        public decimal MaxScore => 2;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var workbook = studentSheet.Workbook;
                var menuSheet = workbook.Worksheets["Menu Items"];
                var documentationSheet = workbook.Worksheets["Documentation"];

                if (menuSheet == null || documentationSheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Menu Items' hoặc 'Documentation'");
                    return result;
                }

                decimal score = 0;
                score += CompareStyle(menuSheet, documentationSheet, "A1", result, 1m);
                score += CompareStyle(menuSheet, documentationSheet, "A2", result, 1m);

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }

        private static decimal CompareStyle(
            ExcelWorksheet menuSheet,
            ExcelWorksheet documentationSheet,
            string address,
            TaskResult result,
            decimal points)
        {
            var menuCell = menuSheet.Cells[address];
            var documentationCell = documentationSheet.Cells[address];

            if (menuCell.StyleID == documentationCell.StyleID)
            {
                result.Details.Add($"{address} có định dạng khớp với Documentation");
                return points;
            }

            result.Errors.Add($"{address} chưa khớp định dạng với Documentation");
            return 0;
        }
    }
}
