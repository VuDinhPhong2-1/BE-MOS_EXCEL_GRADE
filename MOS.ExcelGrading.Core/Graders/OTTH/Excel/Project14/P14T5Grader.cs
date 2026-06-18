using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project14
{
    public class P14T5Grader : ITaskGrader
    {
        public string TaskId => "P14-T5";
        public string TaskName => "Summary: dat Alt Text bieu do = Renewal Data";
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
                var ws = P14GraderHelpers.GetSheet(studentSheet.Workbook, "Summary");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Summary'.");
                    return result;
                }

                var chart = ws.Drawings.FirstOrDefault(d => d is OfficeOpenXml.Drawing.Chart.ExcelChart);
                if (chart == null)
                {
                    result.Errors.Add("Không tìm th?y chart tren sheet 'Summary'.");
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
                    result.Errors.Add($"Alt text chart chua dúng. Hi?n t?i: '{description}'.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




