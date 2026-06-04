using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project11
{
    public sealed class WP11T3Grader : IWordTaskGrader
    {
        private const int ExpectedWidthTwips = 3170;
        private const int WidthToleranceTwips = 60;

        public string TaskId => "W11-T03";
        public string TaskName => "Đặt tất cả cột bảng Filling Agents rộng 5.59 cm";
        public decimal MaxScore => 25m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var table = WP11GraderHelpers.FindFirstTableAfterHeading(studentDocument, "Filling Agents");

            if (table == null)
            {
                WP11GraderHelpers.AddError(
                    result,
                    "Không tìm thấy bảng ngay sau tiêu đề/section “Filling Agents”.",
                    "Khôi phục đúng bảng trong section Filling Agents, chọn toàn bộ bảng rồi đặt Table Layout > Cell Size > Width = 5.59 cm.");
                return result;
            }

            var widths = WP11GraderHelpers.GetTableColumnWidthsTwips(table);
            if (widths.Count == 0)
            {
                WP11GraderHelpers.AddError(
                    result,
                    "Không đọc được độ rộng cột của bảng Filling Agents trong XML.",
                    "Chọn toàn bộ bảng Filling Agents, vào Table Layout > Cell Size và nhập Width = 5.59 cm để Word lưu lại độ rộng cột.");
                return result;
            }

            var correctCount = widths.Count(width => Math.Abs(width - ExpectedWidthTwips) <= WidthToleranceTwips);
            if (correctCount == 0)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Chưa có cột nào của bảng Filling Agents rộng 5.59 cm. Giá trị hiện tại: {string.Join(", ", widths)} twips.",
                    "Chọn toàn bộ bảng trong section Filling Agents, không chỉ chọn một cột, rồi đặt Table Layout > Cell Size > Width = 5.59 cm.");
            }
            else if (correctCount < widths.Count)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Chỉ có {correctCount}/{widths.Count} cột của bảng Filling Agents rộng 5.59 cm; có thể bạn mới chỉnh một vài cột.",
                    "Bôi đen toàn bộ bảng Filling Agents rồi đặt Width = 5.59 cm để tất cả cột có cùng độ rộng.");
            }

            if (result.Errors.Count == 0)
            {
                WP11GraderHelpers.AddDetail(result, $"Tất cả {widths.Count} cột của bảng Filling Agents có độ rộng xấp xỉ 5.59 cm.");
            }

            return result;
        }
    }
}