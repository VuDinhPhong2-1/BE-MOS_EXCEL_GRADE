using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    public class WP03T3Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T3";
        public string TaskName => "Đặt giãn cách dòng của toàn bộ tài liệu thành 1.4 dòng.";
        public decimal MaxScore => 18m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };
            const string fixAction = "Nhấn Ctrl+A để chọn toàn bộ tài liệu, vào Home > Paragraph > Line and Paragraph Spacing > Line Spacing Options, đặt Line spacing là Multiple và At là 1.4.";

            try
            {
                var paragraphs = studentDocument.MainDocumentXml?.Root?
                    .Element(WP03GraderHelpers.W + "body")?
                    .Descendants(WP03GraderHelpers.W + "p")
                    .ToList()
                    ?? new List<XElement>();

                if (paragraphs.Count == 0)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy đoạn văn nào để kiểm tra giãn cách dòng.", "Kiểm tra lại file Word để đảm bảo tài liệu còn nội dung văn bản trước khi đặt giãn cách dòng 1.4.");
                    return result;
                }

                var checkedParagraphs = paragraphs
                    .Where(paragraph =>
                    {
                        var text = WP03GraderHelpers.GetParagraphText(paragraph);
                        var hasDrawing = paragraph.Descendants(WP03GraderHelpers.W + "drawing").Any();
                        return !string.IsNullOrWhiteSpace(text) || hasDrawing;
                    })
                    .ToList();

                if (checkedParagraphs.Count == 0)
                {
                    WP03GraderHelpers.AddError(result, "Không có đoạn văn có nội dung để kiểm tra giãn cách dòng 1.4.", "Kiểm tra lại file Word để đảm bảo tài liệu còn các đoạn văn hoặc đối tượng cần áp dụng giãn cách dòng.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add($"Đã lấy {checkedParagraphs.Count} đoạn để kiểm tra giãn cách dòng.");

                var invalidLineParagraphs = new List<string>();
                var invalidRuleParagraphs = new List<string>();

                foreach (var paragraph in checkedParagraphs)
                {
                    var line = WP03GraderHelpers.GetSpacingLine(paragraph);
                    var lineRule = WP03GraderHelpers.GetSpacingLineRule(paragraph);
                    var sampleText = WP03GraderHelpers.GetParagraphText(paragraph);
                    if (string.IsNullOrWhiteSpace(sampleText))
                    {
                        sampleText = "[Đoạn chứa hình hoặc đối tượng.]";
                    }

                    if (!string.Equals(line, "336", StringComparison.Ordinal))
                    {
                        invalidLineParagraphs.Add($"{sampleText} (line={line}).");
                    }

                    if (!string.Equals(lineRule, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        invalidRuleParagraphs.Add($"{sampleText} (lineRule={lineRule}).");
                    }
                }

                if (invalidLineParagraphs.Count == 0)
                {
                    result.Score += 8m;
                    result.Details.Add("Tất cả đoạn kiểm tra đều có line=336, đúng giãn cách 1.4 dòng.");
                }
                else
                {
                    WP03GraderHelpers.AddError(
                        result,
                        $"Có {invalidLineParagraphs.Count} đoạn chưa có line=336. Ví dụ: {invalidLineParagraphs.First()}",
                        fixAction);
                }

                if (invalidRuleParagraphs.Count == 0)
                {
                    result.Score += 6m;
                    result.Details.Add("Tất cả đoạn kiểm tra đều có lineRule=auto.");
                }
                else
                {
                    WP03GraderHelpers.AddError(
                        result,
                        $"Có {invalidRuleParagraphs.Count} đoạn chưa có lineRule=auto. Ví dụ: {invalidRuleParagraphs.First()}",
                        fixAction);
                }
            }
            catch (Exception ex)
            {
                WP03GraderHelpers.AddError(result, $"Lỗi khi chấm Task 3: {ex.Message}.", "Đóng file Word nếu đang mở, kiểm tra file .docx không bị hỏng rồi tải lại để chấm lại Task 3.");
            }

            return result;
        }
    }
}
