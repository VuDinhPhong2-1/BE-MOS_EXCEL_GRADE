using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project05
{
    public class WP05T6Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T6";
        public string TaskName => "Trong phần \"Woodgrove Essential Savings\", xóa bình luận gắn với dòng chữ \"$3,000\".";
        public decimal MaxScore => 15m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var bodyElements = WP05GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP05GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Woodgrove Essential Savings");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Woodgrove Essential Savings\".");
                    result.FixActions.Add("Kiểm tra lại tiêu đề phần \"Woodgrove Essential Savings\"; không đổi tên hoặc xóa tiêu đề này trước khi xóa comment.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Woodgrove Essential Savings\".");

                var sectionParagraphs = WP05GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: true);
                var targetParagraph = sectionParagraphs.FirstOrDefault(paragraph =>
                    WP05GraderHelpers.GetParagraphText(paragraph).Contains("$3,000", StringComparison.Ordinal));

                if (targetParagraph == null)
                {
                    result.Errors.Add("Không tìm thấy dòng chữ \"$3,000\" để kiểm tra bình luận.");
                    result.FixActions.Add("Tìm trong phần \"Woodgrove Essential Savings\" dòng chứa \"$3,000\" và giữ nguyên định dạng số có dấu phẩy.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy đúng dòng chữ \"$3,000\" với dấu phẩy phân tách hàng nghìn.");

                var hasCommentMarkerOnTarget = targetParagraph.Descendants()
                    .Any(node =>
                        node.Name == WP05GraderHelpers.W + "commentRangeStart"
                        || node.Name == WP05GraderHelpers.W + "commentRangeEnd"
                        || node.Name == WP05GraderHelpers.W + "commentReference");

                if (!hasCommentMarkerOnTarget)
                {
                    result.Score += 5m;
                    result.Details.Add("Đã xóa hoàn toàn comment marker gắn trực tiếp với \"$3,000\".");
                }
                else
                {
                    result.Errors.Add("Vẫn còn comment marker bám trên dòng \"$3,000\".");
                    result.FixActions.Add("Chọn comment gắn với dòng \"$3,000\" và dùng Review > Delete để xóa comment đó khỏi tài liệu.");
                }

                if (studentDocument.TryGetXmlPart("word/comments.xml", out var commentsXml))
                {
                    var commentCount = commentsXml.Root?.Elements(WP05GraderHelpers.W + "comment").Count() ?? 0;
                    if (commentCount <= 1)
                    {
                        result.Score += 3m;
                        result.Details.Add($"Số lượng comment còn lại hợp lệ ({commentCount} comment).");
                    }
                    else
                    {
                        result.Errors.Add($"Tài liệu vẫn còn {commentCount} comment, khả năng chưa xóa đúng comment yêu cầu.");
                        result.FixActions.Add("Mở Review > Comments, kiểm tra danh sách comment và xóa đúng comment được yêu cầu gắn với dòng \"$3,000\".");
                    }
                }
                else
                {
                    result.Score += 3m;
                    result.Details.Add("Không còn part comments.xml, xem như đã xóa toàn bộ comment.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 6: {ex.Message}.");
                result.FixActions.Add("Kiểm tra tài liệu có mở được trong Word và lưu lại ở định dạng .docx trước khi chấm lại.");
            }

            return result;
        }
    }
}
