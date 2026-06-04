using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project05
{
    public class WP05T4Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T4";
        public string TaskName => "Trong đoạn văn trống ngay dưới tiêu đề tài liệu, chèn một mục lục tự động. Sử dụng kiểu \"Automatic Table 1\".";
        public decimal MaxScore => 22m;

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
                var tocNodes = WP05GraderHelpers.GetTocSdtNodes(studentDocument);
                if (tocNodes.Count > 0)
                {
                    result.Score += 8m;
                    result.Details.Add($"Đã chèn mục lục tự động, phát hiện {tocNodes.Count} khối Table of Contents.");
                }
                else
                {
                    result.Errors.Add("Không tìm thấy khối mục lục tự động (Table of Contents).");
                    result.FixActions.Add("Đặt con trỏ ở đoạn trống ngay dưới tiêu đề tài liệu, vào References > Table of Contents và chọn \"Automatic Table 1\".");
                    return result;
                }

                var bodyElements = WP05GraderHelpers.GetBodyElements(studentDocument);
                var titleIndex = WP05GraderHelpers.FindParagraphIndexByExactText(
                    bodyElements,
                    "American Bank Accounts for International Students");

                if (titleIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy dòng tiêu đề chính của tài liệu để kiểm tra vị trí mục lục.");
                    result.FixActions.Add("Khôi phục tiêu đề chính \"American Bank Accounts for International Students\" để mục lục được đặt đúng vị trí yêu cầu.");
                }
                else
                {
                    var firstNodeAfterTitle = bodyElements.Skip(titleIndex + 1).FirstOrDefault();
                    if (firstNodeAfterTitle != null && firstNodeAfterTitle.Name == WP05GraderHelpers.W + "sdt")
                    {
                        result.Score += 5m;
                        result.Details.Add("Mục lục được đặt ngay bên dưới tiêu đề tài liệu.");
                    }
                    else
                    {
                        result.Errors.Add("Mục lục chưa nằm ngay dưới tiêu đề tài liệu.");
                        result.FixActions.Add("Cắt và dán mục lục vào đúng đoạn trống ngay bên dưới tiêu đề chính của tài liệu.");
                    }
                }

                var firstToc = tocNodes[0];
                var fieldInstruction = string.Join(
                    " ",
                    firstToc.Descendants(WP05GraderHelpers.W + "instrText")
                        .Select(node => WP05GraderHelpers.NormalizeText(node.Value)));

                if (fieldInstruction.Contains("TOC \\o \"1-3\" \\h \\z \\u", StringComparison.Ordinal))
                {
                    result.Score += 5m;
                    result.Details.Add("Mục lục dùng đúng trường tự động của kiểu Automatic Table 1.");
                }
                else
                {
                    result.Errors.Add("Khối mục lục chưa đúng cấu hình trường TOC tự động chuẩn.");
                    result.FixActions.Add("Xóa mục lục hiện tại và chèn lại bằng References > Table of Contents > Automatic Table 1 thay vì gõ thủ công hoặc dùng kiểu khác.");
                }

                var tocText = WP05GraderHelpers.NormalizeText(
                    string.Concat(firstToc.Descendants(WP05GraderHelpers.W + "t").Select(node => node.Value)));
                var hasContentsCaption = tocText.Contains("Contents", StringComparison.Ordinal);
                var hasTargetHeading = tocText.Contains("Checking Accounts", StringComparison.Ordinal)
                                       && tocText.Contains("Bank Fees", StringComparison.Ordinal);

                if (hasContentsCaption && hasTargetHeading)
                {
                    result.Score += 4m;
                    result.Details.Add("Mục lục hiển thị đúng tiêu đề và đã lấy được nội dung các mục chính.");
                }
                else
                {
                    result.Errors.Add("Nội dung mục lục chưa đầy đủ hoặc sai chính tả ở các mục chính.");
                    result.FixActions.Add("Cập nhật lại mục lục bằng Update Table và kiểm tra các heading chính như \"Checking Accounts\" và \"Bank Fees\" vẫn đúng chính tả.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 4: {ex.Message}.");
                result.FixActions.Add("Kiểm tra tài liệu có mở được trong Word và lưu lại ở định dạng .docx trước khi chấm lại.");
            }

            return result;
        }
    }
}
