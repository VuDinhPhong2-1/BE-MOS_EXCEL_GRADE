using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project18
{
    public class P18T6Grader : ITaskGrader
    {
        public string TaskId => "P18-T6";
        public string TaskName => "Trong trang tính \"New Accounts\", đối với biểu đồ Account Balances, hoán đổi dữ liệu trên trục để hiển thị Opening Balance và Current Balance dưới dạng chú giải (legend).";
        public decimal MaxScore => 25m;

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
                var worksheet = P18GraderHelpers.GetSheet(studentSheet.Workbook, "New Accounts");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'New Accounts'.");
                    return result;
                }

                decimal score = 0m;

                var chart = worksheet.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Không tìm thấy biểu đồ trên sheet 'New Accounts'.");
                    return result;
                }

                score += 4m;
                result.Details.Add("Đã tìm thấy biểu đồ Account Balances trên sheet 'New Accounts'.");

                if (chart.Series.Count == 2)
                {
                    score += 6m;
                    result.Details.Add("Biểu đồ có đúng 2 series sau khi hoán đổi dữ liệu trục.");
                }
                else
                {
                    result.Errors.Add($"Số series của biểu đồ chưa đúng. Hiện tại: {chart.Series.Count}, mong đợi: 2.");
                }

                var seriesInfos = P18GraderHelpers.ReadChartSeriesInfo(chart);
                var actualHeaders = seriesInfos
                    .Select(info => info.HeaderAddress)
                    .Where(address => !string.IsNullOrWhiteSpace(address))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                var expectedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "B3",
                    "C3"
                };
                var headerMatches = expectedHeaders.Count(header => actualHeaders.Contains(header));
                var headerScore = Math.Round(7m * headerMatches / expectedHeaders.Count, 2, MidpointRounding.AwayFromZero);
                score += headerScore;

                if (headerMatches == expectedHeaders.Count)
                {
                    result.Details.Add("Legend đã lấy đúng 2 tiêu đề cột Opening Balance và Current Balance.");
                }
                else
                {
                    result.Errors.Add(
                        $"Legend chưa hiển thị đúng tiêu đề series mong đợi ({headerMatches}/{expectedHeaders.Count}). Header hiện tại: {string.Join(", ", actualHeaders)}.");
                }

                const string expectedCategoryRange = "A4:A10";
                var xRangeMatches = seriesInfos.Count(info =>
                    string.Equals(info.CategoryAddress, expectedCategoryRange, StringComparison.OrdinalIgnoreCase));
                if (seriesInfos.Count > 0)
                {
                    score += Math.Round(4m * xRangeMatches / seriesInfos.Count, 2, MidpointRounding.AwayFromZero);
                }

                if (seriesInfos.Count > 0 && xRangeMatches == seriesInfos.Count)
                {
                    result.Details.Add("Toàn bộ series đã dùng đúng vùng category A4:A10.");
                }
                else
                {
                    result.Errors.Add(
                        $"Vùng category của biểu đồ chưa đúng cho tất cả series ({xRangeMatches}/{Math.Max(seriesInfos.Count, 1)}).");
                }

                var expectedValueRanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "B4:B10",
                    "C4:C10"
                };
                var actualValueRanges = seriesInfos
                    .Select(info => info.ValueAddress)
                    .Where(address => !string.IsNullOrWhiteSpace(address))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                var valueRangeMatches = expectedValueRanges.Count(expected => actualValueRanges.Contains(expected));
                score += Math.Round(4m * valueRangeMatches / expectedValueRanges.Count, 2, MidpointRounding.AwayFromZero);

                if (valueRangeMatches == expectedValueRanges.Count)
                {
                    result.Details.Add("Vùng dữ liệu series đã đúng: B4:B10 và C4:C10.");
                }
                else
                {
                    result.Errors.Add(
                        $"Vùng dữ liệu series chưa đúng hoàn toàn ({valueRangeMatches}/{expectedValueRanges.Count}). Vùng hiện tại: {string.Join(", ", actualValueRanges)}.");
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


