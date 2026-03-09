using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T6Grader : ITaskGrader
    {
        public string TaskId => "P04-T6";
        public string TaskName => "Dat tieu de truc doc chinh la 'Hours'";
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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Number of course hours");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet Number of course hours");
                    return result;
                }

                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren Number of course hours");
                    return result;
                }

                result.Score += 1m;
                result.Details.Add("Tim thay chart can cham");

                var yTitle = chart.YAxis?.Title?.Text ?? string.Empty;
                if (yTitle.Length > 0)
                {
                    result.Score += 1.5m;
                    result.Details.Add($"Truc doc da co tieu de: '{yTitle}'");
                }
                else
                {
                    result.Errors.Add("Truc doc chua co tieu de");
                }

                if (string.Equals(yTitle, "Hours", StringComparison.Ordinal))
                {
                    result.Score += 1.5m;
                    result.Details.Add("Tieu de truc doc dung 'Hours'");
                }
                else
                {
                    result.Errors.Add($"Tieu de truc doc chua dung chinh ta. Hien tai: '{yTitle}', mong doi dung chinh xac 'Hours'.");
                }

                result.Score = Math.Min(MaxScore, result.Score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
