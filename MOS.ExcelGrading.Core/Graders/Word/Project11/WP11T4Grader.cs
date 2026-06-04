using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project11
{
    public sealed class WP11T4Grader : IWordTaskGrader
    {
        public string TaskId => "W11-T04";
        public string TaskName => "Áp dụng Intense Emphasis cho đoạn ngay bên dưới hình ảnh";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            try
            {
                using var document = WP11GraderHelpers.OpenReadOnlyDocument(studentDocument);
                var paragraphs = WP11GraderHelpers.GetTopLevelParagraphs(document);
                var imageIndex = WP11GraderHelpers.FindMainContentImageParagraphIndex(paragraphs);

                if (imageIndex < 0)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Không xác định được hình ảnh chính trong nội dung tài liệu.",
                        "Chọn đoạn văn ngay dưới hình ảnh chính (đoạn bắt đầu bằng 'This event sẽ diễn ra...'), vào tab Home, trong nhóm Styles, tìm và chọn kiểu Intense Emphasis, rồi lưu file.");
                    return result;
                }

                var targetIndex = WP11GraderHelpers.FindNextMeaningfulParagraphIndex(paragraphs, imageIndex);
                if (targetIndex < 0)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Không tìm thấy đoạn văn ngay bên dưới hình ảnh để kiểm tra style.",
                        "Chọn đoạn văn ngay dưới hình ảnh chính (đoạn bắt đầu bằng 'This event sẽ diễn ra...'), vào tab Home, trong nhóm Styles, tìm và chọn kiểu Intense Emphasis, rồi lưu file.");
                    return result;
                }

                var targetParagraph = paragraphs[targetIndex];
                if (!WP11GraderHelpers.HasStyle(targetParagraph, "IntenseEmphasis"))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Đoạn văn ngay bên dưới hình ảnh chưa được gán style Intense Emphasis.",
                        "Chọn đoạn văn ngay dưới hình ảnh chính (đoạn bắt đầu bằng 'This event sẽ diễn ra...'), vào tab Home, trong nhóm Styles, tìm và chọn kiểu Intense Emphasis, rồi lưu file.");
                }

                var styledNeighbors = paragraphs
                    .Skip(targetIndex + 1)
                    .Where(WP11GraderHelpers.IsMeaningfulParagraph)
                    .Take(2)
                    .Count(paragraph => WP11GraderHelpers.HasStyle(paragraph, "IntenseEmphasis"));

                if (styledNeighbors > 0)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Có đoạn khác gắn vùng hình ảnh cũng đang mang style Intense Emphasis, khả năng áp dụng sai vùng.",
                        "Chọn đoạn văn ngay dưới hình ảnh chính (đoạn bắt đầu bằng 'This event sẽ diễn ra...'), vào tab Home, trong nhóm Styles, tìm và chọn kiểu Intense Emphasis, rồi lưu file.");
                }

                if (result.Errors.Count == 0)
                {
                    WP11GraderHelpers.AddDetail(result, "Doan van ngay sau hinh anh chinh co style Intense Emphasis va khong bi ap dung du sang cac doan ke tiep.");
                }
            }
            catch (Exception ex)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Lỗi khi kiểm tra style dưới hình ảnh: {ex.Message}",
                    "Chọn đoạn văn ngay dưới hình ảnh chính (đoạn bắt đầu bằng 'This event sẽ diễn ra...'), vào tab Home, trong nhóm Styles, tìm và chọn kiểu Intense Emphasis, rồi lưu file.");
            }

            return result;
        }
    }
}
