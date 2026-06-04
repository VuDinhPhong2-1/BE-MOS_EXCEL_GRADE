using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project07
{
    public class WP07T1Grader : IWordTaskGrader
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public string TaskId => "W07-T1";
        public string TaskName => "Lưu bản sao dưới dạng plain text tên Memo";
        public decimal MaxScore => 12m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            var sourceFileName = studentDocument.SourceFileName ?? string.Empty;
            var extension = Path.GetExtension(sourceFileName);
            if (!string.Equals(extension, ".txt", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add("Task này yêu cầu nộp file plain text (.txt).");
                result.FixActions.Add("Mở tài liệu gốc trong Word, chọn File > Save As/Browse, chọn Save as type là Plain Text (*.txt), rồi lưu bản sao.");
                return result;
            }

            result.Score += 6m;
            result.Details.Add("Dùng định dạng plain text (.txt).");

            var baseName = Path.GetFileNameWithoutExtension(sourceFileName) ?? string.Empty;
            if (!string.Equals(baseName, "memo", StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add("Tên file plain text phải là Memo.txt.");
                result.FixActions.Add("Đổi tên file plain text thành Memo.txt trước khi nộp. Tên file phải là Memo và phần mở rộng phải là .txt.");
                return result;
            }

            result.Score += 2m;
            result.Details.Add("Tên file đúng yêu cầu Memo.txt.");

            var documentText = GetDocumentText(studentDocument);
            if (string.IsNullOrWhiteSpace(documentText))
            {
                result.Errors.Add("Không đọc được nội dung plain text để chấm điểm.");
                result.FixActions.Add("Mở Memo.txt để kiểm tra nội dung. Nếu file rỗng hoặc lỗi mã hóa, hãy xuất lại từ tài liệu Word gốc dưới dạng Plain Text (*.txt) và đảm bảo nội dung memo vẫn còn trong file.");
                return result;
            }

            var normalized = NormalizeText(documentText);
            if (normalized.StartsWith("memo", StringComparison.OrdinalIgnoreCase))
            {
                result.Score += 1m;
                result.Details.Add("Nội dung bắt đầu với tiêu đề Memo.");
            }
            else
            {
                result.Errors.Add("Không tìm thấy tiêu đề Memo ở đầu nội dung.");
                result.FixActions.Add("Đặt tiêu đề Memo ở đầu file plain text, trước các dòng To:, From: và CC:.");
            }

            if (ContainsIgnoreCase(normalized, "to:"))
            {
                result.Score += 1m;
            }
            else
            {
                result.Errors.Add("Thiếu dòng To: trong memo.");
                result.FixActions.Add("Bổ sung dòng người nhận bắt đầu bằng To: trong nội dung Memo.txt.");
            }

            if (ContainsIgnoreCase(normalized, "from:"))
            {
                result.Score += 1m;
            }
            else
            {
                result.Errors.Add("Thiếu dòng From: trong memo.");
                result.FixActions.Add("Bổ sung dòng người gửi bắt đầu bằng From: trong nội dung Memo.txt.");
            }

            if (ContainsIgnoreCase(normalized, "cc:"))
            {
                result.Score += 1m;
            }
            else
            {
                result.Errors.Add("Thiếu dòng CC: trong memo.");
                result.FixActions.Add("Bổ sung dòng người nhận bản sao bắt đầu bằng CC: trong nội dung Memo.txt.");
            }

            return result;
        }

        private static string GetDocumentText(WordGradingContext context)
        {
            return string.Concat(
                context.MainDocumentXml?.Descendants(W + "t").Select(node => node.Value) ?? Enumerable.Empty<string>());
        }

        private static string NormalizeText(string value)
        {
            return string.Join(' ', (value ?? string.Empty)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        }

        private static bool ContainsIgnoreCase(string source, string value)
        {
            return source.Contains(value, StringComparison.OrdinalIgnoreCase);
        }
    }
}
