using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project01
{
    public class WP01T2Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T2";
        public string TaskName => "Trong phần \"Children love dinosaurs\", sử dụng công cụ Format Painter để sao chép định dạng của đoạn đầu tiên và áp dụng cho đoạn thứ hai.";
        public decimal MaxScore => 18m;

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
                var bodyElements = WP01GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Children love dinosaurs");
                if (headingIndex < 0)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy tiêu đề \"Children love dinosaurs\".",
                        "Khôi phục đúng tiêu đề \"Children love dinosaurs\" và không đổi nội dung/kiểu Heading của tiêu đề.");
                    return result;
                }

                result.Score += 2m;
                result.Details.Add("Đã tìm thấy đúng phần \"Children love dinosaurs\".");

                var sectionParagraphs = WP01GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: true);
                if (sectionParagraphs.Count < 2)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không đủ 2 đoạn văn để kiểm tra Format Painter trong phần yêu cầu.",
                        "Khôi phục hai đoạn văn bên dưới tiêu đề \"Children love dinosaurs\", sau đó dùng Format Painter từ đoạn 1 sang đoạn 2.");
                    return result;
                }

                var firstParagraph = sectionParagraphs[0];
                var secondParagraph = sectionParagraphs[1];

                var firstText = WP01GraderHelpers.GetParagraphText(firstParagraph);
                var secondText = WP01GraderHelpers.GetParagraphText(secondParagraph);
                if (string.IsNullOrWhiteSpace(firstText) || string.IsNullOrWhiteSpace(secondText))
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Một trong hai đoạn cần kiểm tra đang rỗng nội dung, không thể xác minh định dạng.",
                        "Khôi phục nội dung hai đoạn trong phần \"Children love dinosaurs\" rồi chỉ sao chép định dạng, không xóa văn bản.");
                    return result;
                }

                var firstPPr = WP01GraderHelpers.CanonicalXml(firstParagraph.Element(WP01GraderHelpers.W + "pPr"));
                var secondPPr = WP01GraderHelpers.CanonicalXml(secondParagraph.Element(WP01GraderHelpers.W + "pPr"));
                if (string.Equals(firstPPr, secondPPr, StringComparison.Ordinal))
                {
                    result.Score += 8m;
                    result.Details.Add("Định dạng đoạn (pPr) của đoạn thứ hai đã khớp đoạn đầu tiên.");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn đã không áp dụng định dạng đoạn văn 2 đúng theo đoạn văn 1.",
                        "Chọn đoạn văn đầu tiên, bấm Format Painter, rồi quét đúng đoạn văn thứ hai trong phần \"Children love dinosaurs\".");
                }

                var firstRPr = GetPrimaryRunPropertiesSignature(firstParagraph);
                var secondRPr = GetPrimaryRunPropertiesSignature(secondParagraph);
                if (string.Equals(firstRPr, secondRPr, StringComparison.Ordinal))
                {
                    result.Score += 6m;
                    result.Details.Add("Định dạng ký tự chính (rPr) của đoạn thứ hai đã khớp đoạn đầu tiên.");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn đã lấy nhầm định dạng đoạn văn khác hoặc chưa sao chép đủ định dạng ký tự.",
                        "Dùng Format Painter từ đúng đoạn văn đầu tiên và áp dụng lại cho toàn bộ đoạn văn thứ hai.");
                }

                var hasForbiddenDirectFormat = secondParagraph.Descendants(WP01GraderHelpers.W + "ind").Any()
                    || secondParagraph.Descendants(WP01GraderHelpers.W + "b").Any()
                    || secondParagraph.Descendants(WP01GraderHelpers.W + "i").Any()
                    || secondParagraph.Descendants(WP01GraderHelpers.W + "color").Any()
                    || secondParagraph.Descendants(WP01GraderHelpers.W + "rFonts").Any();

                if (!hasForbiddenDirectFormat)
                {
                    result.Score += 2m;
                    result.Details.Add("Không phát hiện định dạng dư (ind/bold/italic/color/font) trên đoạn thứ hai.");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn đã áp dụng định dạng đoạn văn bị dư hoặc còn định dạng thủ công không đúng trên đoạn văn 2.",
                        "Xóa định dạng thủ công ở đoạn 2 nếu cần, sau đó chỉ dùng Format Painter từ đoạn 1 sang đoạn 2.");
                }
            }
            catch (Exception ex)
            {
                WP01GraderHelpers.AddError(
                    result,
                    $"Lỗi khi chấm Task 2: {ex.Message}.",
                    "Lưu lại tệp .docx và kiểm tra phần \"Children love dinosaurs\" còn đủ hai đoạn văn trước khi chấm lại.");
            }

            return result;
        }

        private static string GetPrimaryRunPropertiesSignature(XElement paragraph)
        {
            var run = paragraph.Elements(WP01GraderHelpers.W + "r")
                .FirstOrDefault(candidate =>
                    candidate.Descendants(WP01GraderHelpers.W + "t").Any(node =>
                        !string.IsNullOrWhiteSpace(node.Value)));

            return WP01GraderHelpers.CanonicalXml(run?.Element(WP01GraderHelpers.W + "rPr"));
        }
    }
}
