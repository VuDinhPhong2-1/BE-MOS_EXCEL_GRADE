using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project20
{
    public class P20T6Grader : ITaskGrader
    {
        public string TaskId => "P20-T6";
        public string TaskName => "Trong trang tính “London”, đối với biểu đồ Air Miles, hiển thị bảng dữ liệu (data table) mà không có ký hiệu chú giải.";
        public decimal MaxScore => 20m;

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
                var worksheet = P20GraderHelpers.GetSheet(studentSheet.Workbook, "London");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'London'.");
                    return result;
                }

                decimal score = 0m;
                var chart = worksheet.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy biểu đồ Air Miles trên sheet 'London'.");
                    return result;
                }

                score += 3m;
                result.Details.Add("Đã tìm thấy biểu đồ Air Miles trên sheet 'London'.");

                if (P20GraderHelpers.TryGetFirstSeriesRanges(chart, out var headerRange, out var categoryRange, out var valueRange)
                    && P20GraderHelpers.IsRangeMatch(headerRange, "D4")
                    && P20GraderHelpers.IsRangeMatch(categoryRange, "B5:B21")
                    && P20GraderHelpers.IsRangeMatch(valueRange, "D5:D21"))
                {
                    score += 5m;
                    result.Details.Add("Biểu đồ vẫn liên kết đúng dữ liệu Air Miles (D4, B5:B21, D5:D21).");
                }
                else
                {
                    result.Errors.Add(
                        $"Dữ liệu nguồn biểu đồ chưa đúng. Header='{headerRange}', Category='{categoryRange}', Value='{valueRange}'.");
                }

                if (P20GraderHelpers.HasDataTable(chart, out var showKeys))
                {
                    if (!showKeys)
                    {
                        score += 7m;
                        result.Details.Add("Biểu đồ đã hiển thị Data Table và tắt ký hiệu chú giải trong Data Table.");
                    }
                    else
                    {
                        result.Errors.Add("Biểu đồ đã có Data Table nhưng vẫn bật ký hiệu chú giải (showKeys).");
                    }
                }
                else
                {
                    result.Errors.Add("Biểu đồ chưa hiển thị Data Table.");
                }

                if (P20GraderHelpers.IsLegendHidden(chart))
                {
                    score += 5m;
                    result.Details.Add("Biểu đồ đã ẩn phần Legend đúng yêu cầu.");
                }
                else
                {
                    result.Errors.Add("Biểu đồ vẫn còn hiển thị Legend.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 6: {ex.Message}.");
            }

            return result;
        }
    }
}


