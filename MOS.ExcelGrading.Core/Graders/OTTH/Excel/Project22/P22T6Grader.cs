using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project22
{
    public class P22T6Grader : ITaskGrader
    {
        public string TaskId => "P22-T6";
        public string TaskName => "Trên trang tính \"Results Distribution\", xóa chú giải khỏi biểu đồ và chỉ hiển thị các giá trị dưới dạng nhãn dữ liệu ở phía trên mỗi cột.";
        public decimal MaxScore => 31m;

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
                decimal score = 0m;

                var chartSheet = P22GraderHelpers.GetChartSheet(studentSheet.Workbook, "Results Distribution");
                if (chartSheet == null)
                {
                    result.Errors.Add("Không tìm thấy chart sheet 'Results Distribution'.");
                    return result;
                }

                score += 4m;
                result.Details.Add("Đã tìm thấy chart sheet 'Results Distribution'.");

                var chart = chartSheet.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy biểu đồ trên chart sheet 'Results Distribution'.");
                    result.Score = score;
                    return result;
                }

                score += 4m;
                result.Details.Add("Chart sheet 'Results Distribution' đã có biểu đồ để chấm.");

                if (P22GraderHelpers.TryGetFirstSeriesRanges(chart, out _, out var categoryRange, out var valueRange)
                    && P22GraderHelpers.IsRangeMatch(categoryRange, "B4:B33")
                    && P22GraderHelpers.IsRangeMatch(valueRange, "G4:G33"))
                {
                    score += 6m;
                    result.Details.Add("Biểu đồ vẫn liên kết đúng dữ liệu Name (B4:B33) và Total Courses (G4:G33).");
                }
                else
                {
                    result.Errors.Add(
                        $"Nguồn dữ liệu của biểu đồ chưa đúng. Category='{categoryRange}', Value='{valueRange}'.");
                }

                if (P22GraderHelpers.IsLegendHidden(chart))
                {
                    score += 7m;
                    result.Details.Add("Legend đã được xóa hoặc ẩn đúng theo yêu cầu.");
                }
                else
                {
                    result.Errors.Add("Legend vẫn còn hiển thị trên biểu đồ.");
                }

                if (P22GraderHelpers.TryGetDataLabelSettings(chart, out var labelSettings))
                {
                    if (labelSettings.ShowValue)
                    {
                        score += 6m;
                        result.Details.Add("Nhãn dữ liệu đã bật hiển thị giá trị trên cột.");
                    }
                    else
                    {
                        result.Errors.Add("Nhãn dữ liệu chưa bật hiển thị giá trị (Show Value).");
                    }

                    if (!labelSettings.ShowLegendKey)
                    {
                        score += 2m;
                        result.Details.Add("Nhãn dữ liệu không hiển thị ký hiệu chú giải (Legend Key).");
                    }
                    else
                    {
                        result.Errors.Add("Nhãn dữ liệu vẫn hiển thị ký hiệu chú giải (Legend Key).");
                    }

                    var normalizedPosition = (labelSettings.Position ?? string.Empty).Trim().ToLowerInvariant();
                    var looksLikeTopPosition = string.IsNullOrWhiteSpace(normalizedPosition)
                                               || normalizedPosition is "outend" or "t" or "bestfit";
                    if (looksLikeTopPosition)
                    {
                        score += 2m;
                        result.Details.Add("Vị trí nhãn dữ liệu phù hợp với hiển thị phía trên cột.");
                    }
                    else
                    {
                        result.Errors.Add($"Vị trí nhãn dữ liệu chưa phù hợp. Hiện tại: '{labelSettings.Position}'.");
                    }

                    if (labelSettings.ShowSeriesName || labelSettings.ShowPercent || labelSettings.ShowBubbleSize)
                    {
                        result.Errors.Add("Nhãn dữ liệu đang hiển thị thêm thông tin không cần thiết.");
                    }
                }
                else
                {
                    result.Errors.Add("Biểu đồ chưa có cấu hình Data Labels để hiển thị giá trị trên cột.");
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

