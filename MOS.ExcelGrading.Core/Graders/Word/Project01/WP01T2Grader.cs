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
                var bodyElements = WP01GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Children love dinosaurs");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Children love dinosaurs\".");
                    return result;
                }

                result.Score += 2m;
                result.Details.Add("Đã tìm thấy đúng phần \"Children love dinosaurs\".");

                var sectionParagraphs = WP01GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: true);
                if (sectionParagraphs.Count < 2)
                {
                    result.Errors.Add("Không đủ 2 đoạn văn để kiểm tra Format Painter trong phần yêu cầu.");
                    return result;
                }

                var firstParagraph = sectionParagraphs[0];
                var secondParagraph = sectionParagraphs[1];

                var firstText = WP01GraderHelpers.GetParagraphText(firstParagraph);
                var secondText = WP01GraderHelpers.GetParagraphText(secondParagraph);
                if (string.IsNullOrWhiteSpace(firstText) || string.IsNullOrWhiteSpace(secondText))
                {
                    result.Errors.Add("Một trong hai đoạn cần kiểm tra đang rỗng nội dung, không thể xác minh định dạng.");
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
                    result.Errors.Add("Định dạng đoạn (pPr) của đoạn thứ hai chưa khớp đoạn đầu tiên.");
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
                    result.Errors.Add("Định dạng ký tự chính (rPr) của đoạn thứ hai chưa khớp đoạn đầu tiên.");
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
                    result.Errors.Add("Đoạn thứ hai còn định dạng dư (thụt lề, đậm, nghiêng, màu, hoặc font) chưa đúng theo yêu cầu Format Painter.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 2: {ex.Message}.");
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
