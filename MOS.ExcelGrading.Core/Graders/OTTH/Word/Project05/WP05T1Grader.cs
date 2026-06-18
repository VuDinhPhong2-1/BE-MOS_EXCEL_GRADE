using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project05
{
    public class WP05T1Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T1";
        public string TaskName => "Hãy tìm từ \"automatic\" và xóa nó khỏi tài liệu.";
        public decimal MaxScore => 12m;

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
                var documentText = WP05GraderHelpers.GetDocumentText(studentDocument);
                if (string.IsNullOrWhiteSpace(documentText))
                {
                    result.Errors.Add("Không đọc được nội dung văn bản để kiểm tra từ \"automatic\".");
                    result.FixActions.Add("Mở lại tài liệu Word và kiểm tra file không bị trống/hỏng trước khi nộp lại.");
                    return result;
                }

                var automaticMatches = Regex.Matches(
                    documentText,
                    @"(?<!\w)automatic(?!\w)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                if (automaticMatches.Count == 0)
                {
                    result.Score += 7m;
                    result.Details.Add("Không còn từ \"automatic\" trong tài liệu.");
                }
                else
                {
                    result.Errors.Add($"Vẫn còn {automaticMatches.Count} từ \"automatic\" trong tài liệu.");
                    result.FixActions.Add("Dùng Find để tìm từ \"automatic\" trong tài liệu, sau đó xóa đúng từ này khỏi câu yêu cầu.");
                }

                const string expectedSentence = "Set up a recurring transfer from your checking account to your savings account.";
                if (documentText.Contains(expectedSentence, StringComparison.Ordinal))
                {
                    result.Score += 5m;
                    result.Details.Add("Câu sau khi xóa từ thừa đúng chính tả, đúng khoảng trắng và đúng dấu chấm câu.");
                }
                else if (documentText.Contains(expectedSentence.TrimEnd('.'), StringComparison.Ordinal))
                {
                    result.Score += 2m;
                    result.Errors.Add("Nội dung thay đổi gần đúng nhưng còn thiếu dấu chấm cuối câu.");
                    result.FixActions.Add("Thêm dấu chấm cuối câu sau câu \"Set up a recurring transfer from your checking account to your savings account.\".");
                }
                else
                {
                    result.Errors.Add("Câu sau khi chỉnh sửa chưa đúng hoàn toàn theo yêu cầu về chữ, khoảng trắng hoặc dấu câu.");
                    result.FixActions.Add("Khôi phục câu thành \"Set up a recurring transfer from your checking account to your savings account.\" với đúng chữ, khoảng trắng và dấu câu.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 1: {ex.Message}.");
                result.FixActions.Add("Kiểm tra tài liệu có mở được trong Word và lưu lại ở định dạng .docx trước khi chấm lại.");
            }

            return result;
        }
    }
}

