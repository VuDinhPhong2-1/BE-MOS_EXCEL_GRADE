using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project04
{
    public class P04T5Grader : ITaskGrader
    {
        public string TaskId => "P04-T5";
        public string TaskName => "Move chart sang chart sheet mới tên Graduation Chart";
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
                var workbook = studentSheet.Workbook;
                var graduationSheet = P04GraderHelpers.GetSheet(studentSheet, "Graduation");
                if (graduationSheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet Graduation");
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
                    result.Details.Add($"Tìm thấy chart sheet '{chartSheet.Name}'");
                }
                else
                {
                    if (chartSheetIgnoreCase != null)
                    {
                        result.Errors.Add($"Tên chart sheet sai: '{chartSheetIgnoreCase.Name}'. Phải dùng chính xác 'Graduation Chart' (phân biệt hoa/thường)");
                    }
                    else
                    {
                        result.Errors.Add("Không tìm thấy chart sheet mới tên Graduation Chart");
                    }
                    return result;
                }

                var chartOnChartSheet = chartSheet.Drawings.OfType<ExcelChart>().Any();
                if (chartOnChartSheet)
                {
                    result.Score += 1m;
                    result.Details.Add("Chart đã được chuyển sang chart sheet mới");
                }
                else
                {
                    result.Errors.Add("Chart sheet mới chưa có chart");
                }

                var graduationStillHasChart = graduationSheet.Drawings.OfType<ExcelChart>().Any();
                if (!graduationStillHasChart)
                {
                    result.Score += 1m;
                    result.Details.Add("Sheet Graduation không còn chart sau khi move");
                }
                else
                {
                    result.Errors.Add("Sheet Graduation vẫn còn chart, chưa move đúng");
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

// minor-sync: non-functional graders update

