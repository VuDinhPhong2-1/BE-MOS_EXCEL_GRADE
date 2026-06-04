using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project05
{
    public class WP05T5Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T5";
        public string TaskName => "Trong phần \"Checking Accounts\", chèn dòng chữ \"24/7 ACCOUNT ACCESS\" vào hộp văn bản màu xanh đậm.";
        public decimal MaxScore => 16m;

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
                var headingIndex = WP05GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Checking Accounts");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Checking Accounts\".");
                    result.FixActions.Add("Kiểm tra lại tiêu đề phần \"Checking Accounts\"; không đổi tên hoặc xóa tiêu đề này trước khi chèn chữ vào hộp văn bản.");
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
                    result.FixActions.Add("Trong phần \"Checking Accounts\", chọn hộp văn bản màu xanh đậm có sẵn và nhập dòng \"24/7 ACCOUNT ACCESS\" vào bên trong.");
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
                    result.FixActions.Add("Sửa nội dung trong hộp văn bản màu xanh đậm thành chính xác \"24/7 ACCOUNT ACCESS\" bằng chữ hoa, đúng dấu / và không thêm chữ khác.");
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
                    result.FixActions.Add("Nhập chính xác \"24/7 ACCOUNT ACCESS\" vào hộp văn bản màu xanh đậm trong phần \"Checking Accounts\".");
                }
                else
                {
                    result.Errors.Add($"Có {exactCount} hộp văn bản cùng chứa cụm chữ yêu cầu, cần kiểm tra lại để tránh chèn dư.");
                    result.FixActions.Add("Giữ cụm \"24/7 ACCOUNT ACCESS\" chỉ trong đúng một hộp văn bản màu xanh đậm và xóa các bản chèn dư ở hộp văn bản khác.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 5: {ex.Message}.");
                result.FixActions.Add("Kiểm tra tài liệu có mở được trong Word và lưu lại ở định dạng .docx trước khi chấm lại.");
            }

            return result;
        }
    }
}
