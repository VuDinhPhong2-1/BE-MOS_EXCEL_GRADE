using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project14
{
    public class P14T5Grader : ITaskGrader
    {
        public string TaskId => "P14-T5";
        public string TaskName => "Summary: dat Alt Text bieu do = Renewal Data";
        public decimal MaxScore => 4;

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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Summary'.");
                    return result;
                }

                var chart = ws.Drawings.FirstOrDefault(d => d is OfficeOpenXml.Drawing.Chart.ExcelChart);
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy chart tren sheet 'Summary'.");
                    return result;
                }

                var description = chart.GetType().GetProperty("Description")?.GetValue(chart)?.ToString() ?? string.Empty;
                if (string.Equals(description.Trim(), "Renewal Data", StringComparison.Ordinal))
                {
                    result.Score = MaxScore;
                    result.Details.Add("Alt text chart dung: 'Renewal Data'.");
                }
                else
                {
                    result.Errors.Add($"Alt text chart chưa đúng. Hiện tại: '{description}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}



