using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project05
{
    public class WP05T5Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T5";
        public string TaskName => "Trong phần \"Checking Accounts\", chèn dòng chữ \"24/7 ACCOUNT ACCESS\" vào hộp văn bản màu xanh đậm.";
        public decimal MaxScore => 16m;

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
                var bodyElements = WP05GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP05GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Checking Accounts");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Checking Accounts\".");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Checking Accounts\".");

                var sectionElements = WP05GraderHelpers.GetSectionElements(bodyElements, headingIndex, stopAtHeading1: true);
                // Ưu tiên nhánh DrawingML (wps:txbx). Nhánh VML fallback có thể lặp nội dung và gây đếm trùng.
                var textboxNodes = sectionElements
                    .SelectMany(element => element
                        .Descendants(WP05GraderHelpers.Wps + "txbx")
                        .SelectMany(node => node.Elements(WP05GraderHelpers.W + "txbxContent")))
                    .ToList();

                if (textboxNodes.Count == 0)
                {
                    // Fallback khi tài liệu chỉ có VML textbox.
                    textboxNodes = sectionElements
                        .SelectMany(element => element
                            .Descendants()
                            .Where(node => string.Equals(node.Name.LocalName, "textbox", StringComparison.OrdinalIgnoreCase))
                            .SelectMany(node => node.Descendants(WP05GraderHelpers.W + "txbxContent")))
                        .ToList();
                }

                if (textboxNodes.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy hộp văn bản trong phần \"Checking Accounts\".");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add($"Đã tìm thấy {textboxNodes.Count} hộp văn bản để kiểm tra nội dung.");

                const string expectedText = "24/7 ACCOUNT ACCESS";
                var normalizedTexts = textboxNodes
                    .Select(WP05GraderHelpers.GetTextboxText)
                    .Where(text => !string.IsNullOrWhiteSpace(text))
                    .ToList();

                if (normalizedTexts.Any(text => string.Equals(text, expectedText, StringComparison.Ordinal)))
                {
                    result.Score += 6m;
                    result.Details.Add("Nội dung hộp văn bản đúng chính xác \"24/7 ACCOUNT ACCESS\".");
                }
                else
                {
                    var sample = normalizedTexts.FirstOrDefault() ?? "[Rỗng]";
                    result.Errors.Add($"Nội dung hộp văn bản chưa đúng. Giá trị hiện tại: \"{sample}\".");
                }

                var exactCount = normalizedTexts.Count(text => string.Equals(text, expectedText, StringComparison.Ordinal));
                if (exactCount == 1)
                {
                    result.Score += 3m;
                    result.Details.Add("Không có nội dung dư hoặc sai biến thể trong các hộp văn bản liên quan.");
                }
                else if (exactCount == 0)
                {
                    result.Errors.Add("Không có hộp văn bản nào chứa đúng cụm chữ yêu cầu.");
                }
                else
                {
                    result.Errors.Add($"Có {exactCount} hộp văn bản cùng chứa cụm chữ yêu cầu, cần kiểm tra lại để tránh chèn dư.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 5: {ex.Message}.");
            }

            return result;
        }
    }
}
