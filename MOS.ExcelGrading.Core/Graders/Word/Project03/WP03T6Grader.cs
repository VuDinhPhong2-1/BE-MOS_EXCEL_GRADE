using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    public class WP03T6Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T6";
        public string TaskName => "Trong phần \"Serving\", thay đổi cách ngắt dòng văn bản cho hình ảnh thành Square.";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument, WordGradingContext? answerDocument = null)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var bodyElements = WP03GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Serving");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Serving\".");
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
                    result.Errors.Add("Không tìm thấy hình ảnh trong phần \"Serving\".");
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
                    result.Errors.Add("Hình ảnh vẫn đang ở chế độ In Line with Text (wp:inline), chưa đổi sang kiểu wrap Square.");
                    return result;
                }
                else
                {
                    result.Errors.Add("Không nhận diện được container anchor/inline của hình ảnh.");
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
                    result.Errors.Add($"Kiểu ngắt dòng hiện tại chưa phải Square (đang là {actualWrap}).");
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
                    result.Errors.Add("Đoạn văn cạnh hình ảnh thiếu dấu câu kết thúc hoặc sai dấu chấm câu.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 6: {ex.Message}.");
            }

            return result;
        }
    }
}
