using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project14
{
    public class P14T2Grader : ITaskGrader
    {
        public string TaskId => "P14-T2";
        public string TaskName => "March: loc Policy Type = PM";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "March");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'March'.");
                    return result;
                }

                decimal score = 0m;
                var table = P14GraderHelpers.FindTableByAddress(ws, "A4:G24");
                if (table != null)
                {
                    score += 1m;
                    result.Details.Add("Tìm thấy table March A4:G24.");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy table March A4:G24.");
                    result.Score = score;
                    return result;
                }

                var filterValue = P14GraderHelpers.GetSingleFilterValue(table, 5);
                if (string.Equals(filterValue, "PM", StringComparison.Ordinal))
                {
                    score += 3m;
                    result.Details.Add("Filter cột Policy Type dung giá trị 'PM'.");
                }
                else
                {
                    result.Errors.Add($"Filter Policy Type chưa đúng. Hiện tại: '{filterValue}'.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}



