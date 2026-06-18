using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project03
{
    public class WP03T6Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T6";
        public string TaskName => "Trong phần \"Serving\", thay đổi cách ngắt dòng văn bản cho hình ảnh thành Square.";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };
            const string fixAction = "Trong phần Serving, chọn hình ảnh, vào Picture Format > Wrap Text và chọn Square để văn bản bao quanh hình theo dạng hình vuông.";

            try
            {
                var bodyElements = WP03GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Serving");
                if (headingIndex < 0)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy tiêu đề \"Serving\".", "Kiểm tra lại tài liệu và đảm bảo vẫn còn tiêu đề \"Serving\" đúng chính tả trước khi chỉnh Wrap Text cho hình ảnh.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Serving\".");

                var sectionElements = WP03GraderHelpers.GetSectionElements(bodyElements, headingIndex, stopAtHeading1: false);
                var firstDrawing = sectionElements
                    .SelectMany(element => element.Descendants(WP03GraderHelpers.W + "drawing"))
                    .FirstOrDefault();

                if (firstDrawing == null)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy hình ảnh trong phần \"Serving\".", "Khôi phục hoặc chèn lại hình ảnh trong phần Serving, sau đó chọn hình và đặt Wrap Text là Square.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy hình ảnh cần kiểm tra kiểu ngắt dòng.");

                var anchorNode = firstDrawing.Element(WP03GraderHelpers.WP + "anchor");
                var inlineNode = firstDrawing.Element(WP03GraderHelpers.WP + "inline");

                if (anchorNode != null)
                {
                    result.Score += 5m;
                    result.Details.Add("Hình ảnh đang ở chế độ floating (wp:anchor), phù hợp để đặt wrap.");
                }
                else if (inlineNode != null)
                {
                    WP03GraderHelpers.AddError(result, "Hình ảnh vẫn đang ở chế độ In Line with Text (wp:inline), chưa đổi sang kiểu wrap Square.", fixAction);
                    return result;
                }
                else
                {
                    WP03GraderHelpers.AddError(result, "Không nhận diện được container anchor/inline của hình ảnh.", "Chọn lại đúng hình ảnh trong phần Serving, áp dụng Wrap Text > Square rồi lưu lại file .docx.");
                    return result;
                }

                var wrapSquare = anchorNode!.Element(WP03GraderHelpers.WP + "wrapSquare");
                var wrapText = wrapSquare?.Attribute("wrapText")?.Value ?? string.Empty;
                if (wrapSquare != null
                    && (string.IsNullOrWhiteSpace(wrapText)
                        || string.Equals(wrapText, "bothSides", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(wrapText, "largest", StringComparison.OrdinalIgnoreCase)))
                {
                    result.Score += 6m;
                    result.Details.Add("Hình ảnh đã được đặt kiểu ngắt dòng Square.");
                }
                else
                {
                    var actualWrap = anchorNode.Elements().FirstOrDefault(node => node.Name.LocalName.StartsWith("wrap", StringComparison.OrdinalIgnoreCase))?.Name.LocalName
                                     ?? "không xác định";
                    WP03GraderHelpers.AddError(result, $"Kiểu ngắt dòng hiện tại chưa phải Square (đang là {actualWrap}).", fixAction);
                }

                var paragraph = firstDrawing.Ancestors(WP03GraderHelpers.W + "p").FirstOrDefault();
                var paragraphText = paragraph == null ? string.Empty : WP03GraderHelpers.GetParagraphText(paragraph);
                if (paragraphText.EndsWith(".", StringComparison.Ordinal))
                {
                    result.Score += 2m;
                    result.Details.Add("Đoạn văn cạnh hình ảnh giữ dấu câu kết thúc đầy đủ.");
                }
                else
                {
                    WP03GraderHelpers.AddError(result, "Đoạn văn cạnh hình ảnh thiếu dấu câu kết thúc hoặc sai dấu chấm câu.", "Kiểm tra đoạn văn cạnh hình ảnh trong phần Serving và khôi phục dấu chấm kết thúc câu nếu đã bị xóa hoặc sửa sai.");
                }
            }
            catch (Exception ex)
            {
                WP03GraderHelpers.AddError(result, $"Lỗi khi chấm Task 6: {ex.Message}.", "Đóng file Word nếu đang mở, kiểm tra file .docx không bị hỏng rồi tải lại để chấm lại Task 6.");
            }

            return result;
        }
    }
}

