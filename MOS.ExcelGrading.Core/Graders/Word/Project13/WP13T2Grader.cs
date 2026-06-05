using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T2Grader : IWordTaskGrader
    {
        public WP13T2Grader(string taskId = "W13-T02")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Đặt trang 3 ở hướng Landscape";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var sectionProperties = WP13GraderHelpers.GetSectionPropertiesInDocumentOrder(studentDocument);

            if (sectionProperties.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Không tìm thấy section properties để xác định hướng trang.",
                    "Đặt con trỏ ở trang 3, vào Layout > Breaks để tách section nếu cần, sau đó chọn Orientation > Landscape cho riêng trang 3.");
                return result;
            }

            var landscapeSections = sectionProperties
                .Select((section, index) => new { Section = section, Index = index })
                .Where(item => WP13GraderHelpers.IsLandscapeSection(item.Section))
                .ToList();

            if (landscapeSections.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Chưa phát hiện section nào được đặt hướng Landscape cho trang 3.",
                    "Đặt con trỏ ở trang 3, rồi vào Layout > Orientation > Landscape.");
            }
            else if (landscapeSections.Count == sectionProperties.Count && sectionProperties.Count > 1)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Toàn bộ tài liệu hoặc quá nhiều section đang ở hướng Landscape, chưa giới hạn riêng trang 3.",
                    "Chỉ chọn/tách riêng trang 3 rồi áp dụng Layout > Orientation > Landscape; các trang còn lại giữ Portrait.");
            }
            else if (landscapeSections.Count > 1)
            {
                WP13GraderHelpers.AddError(
                    result,
                    $"Phát hiện {landscapeSections.Count} section Landscape; yêu cầu chỉ áp dụng cho trang 3.",
                    "Kiểm tra lại section breaks quanh trang 3 và đổi các section không thuộc trang 3 về Portrait.");
            }

            if (sectionProperties.Count >= 3)
            {
                var thirdSection = sectionProperties[2];
                if (!WP13GraderHelpers.IsLandscapeSection(thirdSection))
                {
                    WP13GraderHelpers.AddError(
                        result,
                        "Section tương ứng vị trí trang/section thứ 3 chưa ở hướng Landscape.",
                        "Đặt con trỏ ở trang 3, rồi vào Layout > Orientation > Landscape.");
                }
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "Tài liệu có section Landscape hợp lệ để biểu diễn trang 3 theo hướng ngang.");
            }

            return result;
        }
    }
}
