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
                var paragraphs = studentDocument.MainDocumentXml?.Root?
                    .Element(WP03GraderHelpers.W + "body")?
                    .Descendants(WP03GraderHelpers.W + "p")
                    .ToList()
                    ?? new List<XElement>();

                if (paragraphs.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy đoạn văn nào để kiểm tra giãn cách dòng.");
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
                    result.Errors.Add("Không có đoạn văn có nội dung để kiểm tra giãn cách dòng 1.4.");
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
                    result.Errors.Add(
                        $"Có {invalidLineParagraphs.Count} đoạn chưa có line=336. Ví dụ: {invalidLineParagraphs.First()}");
                }

                if (invalidRuleParagraphs.Count == 0)
                {
                    result.Score += 6m;
                    result.Details.Add("Tất cả đoạn kiểm tra đều có lineRule=auto.");
                }
                else
                {
                    result.Errors.Add(
                        $"Có {invalidRuleParagraphs.Count} đoạn chưa có lineRule=auto. Ví dụ: {invalidRuleParagraphs.First()}");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 3: {ex.Message}.");
            }

            return result;
        }
    }
}
