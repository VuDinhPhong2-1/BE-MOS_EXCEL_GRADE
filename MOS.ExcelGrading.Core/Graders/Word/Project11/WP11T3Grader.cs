using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project11
{
    public sealed class WP11T3Grader : IWordTaskGrader
    {
        private const int ExpectedLineTwips = 280;

        public string TaskId => "W11-T03";
        public string TaskName => "Đặt giãn dòng 14 pt cho hai đoạn cuối tài liệu";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            try
            {
                using var document = WP11GraderHelpers.OpenReadOnlyDocument(studentDocument);
                var nonEmptyParagraphs = WP11GraderHelpers.GetMeaningfulParagraphs(document)
                    .Where(paragraph => !string.IsNullOrWhiteSpace(WP11GraderHelpers.GetParagraphText(paragraph)))
                    .ToList();

                if (nonEmptyParagraphs.Count < 2)
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Không đủ đoạn văn để xác định hai đoạn cuối tài liệu.",
                        "Chọn đúng hai đoạn văn cuối cùng ở cuối tài liệu, click chuột phải chọn Paragraph (hoặc vào tab Home > nhóm Paragraph, click mở rộng settings), ở mục Spacing > Line spacing chọn Exactly, tại ô At nhập 14 pt, rồi click OK và lưu file.");
                    return result;
                }

                var lastTwoParagraphs = nonEmptyParagraphs.TakeLast(2).ToList();
                if (!lastTwoParagraphs.All(paragraph => WP11GraderHelpers.HasExactLineSpacing(paragraph, ExpectedLineTwips)))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Hai đoạn văn cuối chưa có w:spacing/@w:line = 280 và w:lineRule = exact.",
                        "Chọn đúng hai đoạn văn cuối cùng ở cuối tài liệu, click chuột phải chọn Paragraph (hoặc vào tab Home > nhóm Paragraph, click mở rộng settings), ở mục Spacing > Line spacing chọn Exactly, tại ô At nhập 14 pt, rồi click OK và lưu file.");
                }

                if (nonEmptyParagraphs.Count >= 3
                    && WP11GraderHelpers.HasExactLineSpacing(nonEmptyParagraphs[^3], ExpectedLineTwips))
                {
                    WP11GraderHelpers.AddError(
                        result,
                        "Có dấu hiệu đã áp dụng giãn dòng 14 pt cho hơn hai đoạn ở cuối tài liệu.",
                        "Chọn đúng hai đoạn văn cuối cùng ở cuối tài liệu, click chuột phải chọn Paragraph (hoặc vào tab Home > nhóm Paragraph, click mở rộng settings), ở mục Spacing > Line spacing chọn Exactly, tại ô At nhập 14 pt, rồi click OK và lưu file.");
                }

                if (result.Errors.Count == 0)
                {
                    WP11GraderHelpers.AddDetail(result, "Hai đoạn văn cuối có line spacing exact 14 pt (w:line=280) và đoạn đứng trước đó không bị áp dụng nhầm.");
                }
            }
            catch (Exception ex)
            {
                WP11GraderHelpers.AddError(
                    result,
                    $"Lỗi khi kiểm tra giãn dòng hai đoạn cuối: {ex.Message}",
                    "Chọn đúng hai đoạn văn cuối cùng ở cuối tài liệu, click chuột phải chọn Paragraph (hoặc vào tab Home > nhóm Paragraph, click mở rộng settings), ở mục Spacing > Line spacing chọn Exactly, tại ô At nhập 14 pt, rồi click OK và lưu file.");
            }

            return result;
        }
    }
}
