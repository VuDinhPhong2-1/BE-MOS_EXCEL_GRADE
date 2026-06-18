using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project20
{
    public class P20T5Grader : ITaskGrader
    {
        public string TaskId => "P20-T5";
        public string TaskName => "Trong trang tính “New York City”, tạo biểu đồ cột dạng Clustered Column để hiển thị Air Miles của tất cả các thành phố, với các thành phố là nhãn trên trục ngang (trục X). Đặt biểu đồ bên dưới bảng. Kích thước và vị trí chính xác như yêu cầu.";
        public decimal MaxScore => 28m;

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
                var worksheet = P20GraderHelpers.GetSheet(studentSheet.Workbook, "New York City");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'New York City'.");
                    return result;
                }

                decimal score = 0m;
                var chart = worksheet.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy biểu đồ trên sheet 'New York City'.");
                    return result;
                }

                score += 4m;
                result.Details.Add("Đã tìm thấy biểu đồ trên sheet 'New York City'.");

                if (P20GraderHelpers.IsClusteredColumnChart(chart))
                {
                    score += 6m;
                    result.Details.Add("Biểu đồ đúng loại Clustered Column.");
                }
                else
                {
                    result.Errors.Add($"Biểu đồ chưa đúng loại Clustered Column. Loại hiện tại: {chart.ChartType}.");
                }

                var seriesCount = chart.Series.Count;
                var seriesCountScore = seriesCount == 1 ? 2m : 0m;
                score += seriesCountScore;
                if (seriesCount == 1)
                {
                    result.Details.Add("Biểu đồ có đúng 1 series dữ liệu.");
                }
                else
                {
                    result.Errors.Add($"Biểu đồ có số series chưa đúng. Hiện tại: {seriesCount}, mong đợi: 1.");
                }

                if (P20GraderHelpers.TryGetFirstSeriesRanges(chart, out var headerRange, out var categoryRange, out var valueRange))
                {
                    if (P20GraderHelpers.IsRangeMatch(categoryRange, "B5:B21"))
                    {
                        score += 4m;
                        result.Details.Add("Nhãn trục X lấy đúng từ cột City (B5:B21).");
                    }
                    else
                    {
                        result.Errors.Add($"Nhãn trục X chưa đúng. Hiện tại: '{categoryRange}', mong đợi: 'B5:B21'.");
                    }

                    if (P20GraderHelpers.IsRangeMatch(valueRange, "D5:D21"))
                    {
                        score += 4m;
                        result.Details.Add("Dữ liệu series lấy đúng từ cột Air Miles (D5:D21).");
                    }
                    else
                    {
                        result.Errors.Add($"Vùng dữ liệu series chưa đúng. Hiện tại: '{valueRange}', mong đợi: 'D5:D21'.");
                    }

                    if (P20GraderHelpers.IsRangeMatch(headerRange, "D4"))
                    {
                        score += 2m;
                        result.Details.Add("Tiêu đề series lấy đúng từ ô D4 (Air Miles).");
                    }
                    else
                    {
                        result.Errors.Add($"Tiêu đề series chưa đúng. Hiện tại: '{headerRange}', mong đợi: 'D4'.");
                    }
                }
                else
                {
                    result.Errors.Add("Không đọc được thông tin series của biểu đồ.");
                }

                if (P20GraderHelpers.IsChartTitleBlank(chart))
                {
                    score += 4m;
                    result.Details.Add("Biểu đồ đã bỏ tiêu đề đúng yêu cầu.");
                }
                else
                {
                    result.Errors.Add("Biểu đồ chưa bỏ tiêu đề.");
                }

                var actualBounds = P20GraderHelpers.GetChartBounds(chart);
                if (P20GraderHelpers.IsRangeMatch(actualBounds, "A25:D39"))
                {
                    score += 4m;
                    result.Details.Add("Vị trí và kích thước biểu đồ đúng vùng A25:D39 (bên dưới bảng).");
                }
                else
                {
                    result.Errors.Add(
                        $"Vị trí hoặc kích thước biểu đồ chưa đúng. Hiện tại: '{actualBounds}', mong đợi: 'A25:D39'.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 5: {ex.Message}.");
            }

            return result;
        }
    }
}


