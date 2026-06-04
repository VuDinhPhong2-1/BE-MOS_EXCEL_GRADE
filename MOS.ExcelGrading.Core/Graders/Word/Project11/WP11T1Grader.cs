using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project11
{
    public sealed class WP11T1Grader : IWordTaskGrader
    {
        public string TaskId => "W11-T01";
        public string TaskName => "Thêm đường viền trang Box màu Dark Blue, Accent 1, dày 3 pt";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            try
            {
                using var document = WP11GraderHelpers.OpenReadOnlyDocument(studentDocument);
                var sections = WP11GraderHelpers.GetSectionProperties(document);
                if (sections.Count == 0)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Không tìm thấy SectionProperties để kiểm tra đường viền trang.",
                        "Trong tab Design, chọn Page Borders; ở mục Setting chọn kiểu Box; ở mục Color chọn màu Dark Blue, Accent 1; ở mục Width chọn 3 pt; Apply to chọn Whole document; rồi nhấn OK và lưu file.");
                    return result;
                }

                var invalidSections = sections.Count(section => !WP11GraderHelpers.SectionHasExpectedPageBorders(section));
                if (invalidSections > 0)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        $"Có {invalidSections}/{sections.Count} section chưa có đủ đường viền trang đúng kiểu Box, màu Accent 1 và độ dày 3 pt.",
                        "Trong tab Design, chọn Page Borders; ở mục Setting chọn kiểu Box; ở mục Color chọn màu Dark Blue, Accent 1; ở mục Width chọn 3 pt; Apply to chọn Whole document; rồi nhấn OK và lưu file.");
                    return result;
                }

                WP11GraderHelpers.AddDetail(result, $"Tất cả {sections.Count} section đều có w:pgBorders với 4 cạnh, w:val=\"single\", w:sz=\"24\" và màu Accent 1.");
            }
            catch (Exception ex)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Lỗi khi kiểm tra đường viền trang: {ex.Message}",
                    "Trong tab Design, chọn Page Borders; ở mục Setting chọn kiểu Box; ở mục Color chọn màu Dark Blue, Accent 1; ở mục Width chọn 3 pt; Apply to chọn Whole document; rồi nhấn OK và lưu file.");
            }

            return result;
        }
    }
}
