using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T5Grader : ITaskGrader
    {
        public string TaskId => "P04-T5";
        public string TaskName => "Move chart sang chart sheet moi ten Graduation Chart";
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
                var workbook = studentSheet.Workbook;
                var graduationSheet = P04GraderHelpers.GetSheet(studentSheet, "Graduation");
                if (graduationSheet == null)
                {
                    result.Errors.Add("Khong tim thay sheet Graduation");
                    return result;
                }

                var chartSheet = workbook.Worksheets
                    .FirstOrDefault(w => w is ExcelChartsheet && P04GraderHelpers.IsGraduationChartSheetName(w.Name));

                var chartSheetIgnoreCase = workbook.Worksheets
                    .FirstOrDefault(w =>
                        w is ExcelChartsheet &&
                        string.Equals((w.Name ?? string.Empty).Trim(), "Graduation Chart", StringComparison.OrdinalIgnoreCase));

                if (chartSheet != null)
                {
                    result.Score += 2m;
                    result.Details.Add($"Tim thay chart sheet '{chartSheet.Name}'");
                }
                else
                {
                    if (chartSheetIgnoreCase != null)
                    {
                        result.Errors.Add($"Ten chart sheet sai: '{chartSheetIgnoreCase.Name}'. Phai dung chinh xac 'Graduation Chart' (phan biet hoa/thuong)");
                    }
                    else
                    {
                        result.Errors.Add("Khong tim thay chart sheet moi ten Graduation Chart");
                    }
                    return result;
                }

                var chartOnChartSheet = chartSheet.Drawings.OfType<ExcelChart>().Any();
                if (chartOnChartSheet)
                {
                    result.Score += 1m;
                    result.Details.Add("Chart da duoc chuyen sang chart sheet moi");
                }
                else
                {
                    result.Errors.Add("Chart sheet moi chua co chart");
                }

                var graduationStillHasChart = graduationSheet.Drawings.OfType<ExcelChart>().Any();
                if (!graduationStillHasChart)
                {
                    result.Score += 1m;
                    result.Details.Add("Sheet Graduation khong con chart sau khi move");
                }
                else
                {
                    result.Errors.Add("Sheet Graduation van con chart, chua move dung");
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
