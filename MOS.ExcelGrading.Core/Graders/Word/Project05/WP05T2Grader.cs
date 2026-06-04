using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project05
{
    public class WP05T2Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T2";
        public string TaskName => "Sử dụng tính năng trong Word để thay thế tất cả các cụm từ \"Woodgrove Savings\" bằng \"Woodgrove Plus\".";
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
                var documentText = WP05GraderHelpers.GetDocumentText(studentDocument);
                if (string.IsNullOrWhiteSpace(documentText))
                {
                    result.Errors.Add("Không đọc được nội dung tài liệu để kiểm tra thao tác thay thế.");
                    result.FixActions.Add("Mở lại tài liệu Word, kiểm tra nội dung không bị trống/hỏng và lưu lại ở định dạng .docx.");
                    return result;
                }

                var oldPhraseCount = WP05GraderHelpers.CountExactPhrase(documentText, "Woodgrove Savings", ignoreCase: true);
                if (oldPhraseCount == 0)
                {
                    result.Score += 4m;
                    result.Details.Add("Không còn cụm \"Woodgrove Savings\" trong tài liệu.");
                }
                else
                {
                    result.Errors.Add($"Vẫn còn {oldPhraseCount} cụm \"Woodgrove Savings\" chưa được thay thế.");
                    result.FixActions.Add("Dùng Home > Replace để thay thế tất cả \"Woodgrove Savings\" bằng \"Woodgrove Plus\" trong toàn bộ tài liệu.");
                }

                var newPhraseExactCount = WP05GraderHelpers.CountExactPhrase(documentText, "Woodgrove Plus", ignoreCase: false);
                if (newPhraseExactCount >= 3)
                {
                    result.Score += 6m;
                    result.Details.Add($"Đã tìm thấy {newPhraseExactCount} cụm \"Woodgrove Plus\" đúng chính tả.");
                }
                else
                {
                    result.Errors.Add($"Số lần xuất hiện \"Woodgrove Plus\" chưa đủ. Hiện tại chỉ có {newPhraseExactCount} lần.");
                    result.FixActions.Add("Chạy Replace All lại với Find what: \"Woodgrove Savings\" và Replace with: \"Woodgrove Plus\" để cập nhật đủ mọi vị trí.");
                }

                var newPhraseIgnoreCaseCount = WP05GraderHelpers.CountExactPhrase(documentText, "Woodgrove Plus", ignoreCase: true);
                if (newPhraseIgnoreCaseCount == newPhraseExactCount)
                {
                    result.Score += 4m;
                    result.Details.Add("Cụm thay thế đúng chữ hoa, chữ thường và không sai biến thể.");
                }
                else
                {
                    result.Errors.Add("Có cụm thay thế sai kiểu chữ hoa, chữ thường hoặc sai chính tả gần đúng.");
                    result.FixActions.Add("Kiểm tra lại các cụm đã thay thế và sửa chính xác thành \"Woodgrove Plus\" đúng chữ hoa/chữ thường.");
                }

                var bodyElements = WP05GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP05GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Woodgrove Plus");
                if (headingIndex >= 0)
                {
                    result.Score += 4m;
                    result.Details.Add("Đã cập nhật đúng tiêu đề mục thành \"Woodgrove Plus\".");
                }
                else
                {
                    result.Errors.Add("Chưa đổi đúng tiêu đề mục thành \"Woodgrove Plus\".");
                    result.FixActions.Add("Tìm tiêu đề mục cũ và đổi đúng nội dung tiêu đề thành \"Woodgrove Plus\".");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 2: {ex.Message}.");
                result.FixActions.Add("Kiểm tra tài liệu có mở được trong Word và lưu lại ở định dạng .docx trước khi chấm lại.");
            }

            return result;
        }
    }
}
