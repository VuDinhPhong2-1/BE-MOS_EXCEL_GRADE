using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T1Grader : IWordTaskGrader
    {
        public WP13T1Grader(string taskId = "W13-T01")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Kiểm tra Accessibility và thêm tiêu đề cho bảng";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var tables = WP13GraderHelpers.GetTables(studentDocument);

            if (tables.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Không tìm thấy bảng trong tài liệu để kiểm tra lỗi Accessibility về Table Title.",
                    "Khôi phục đúng bảng trong tài liệu, sau đó vào Review > Check Accessibility, chọn lỗi liên quan đến bảng và áp dụng recommended action đầu tiên để thêm Table Title.");
                return result;
            }

            var tablesWithoutTitle = tables
                .Where(table => !WP13GraderHelpers.HasTableTitleOrCaption(table))
                .ToList();

            if (tablesWithoutTitle.Count > 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    $"Có {tablesWithoutTitle.Count} bảng chưa có tiêu đề/caption dùng cho Accessibility.",
                    "Vào Review > Check Accessibility, mở lỗi Table Title, chọn recommended action đầu tiên và nhập tiêu đề phù hợp cho bảng.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "Tất cả bảng có dấu hiệu Table Title/caption hoặc metadata mô tả dùng cho Accessibility.");
            }

            return result;
        }
    }
}
