using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    public class WP03T1Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T1";
        public string TaskName => "Hiển thị tiêu đề \"Integral\" trên tất cả các trang của tài liệu, trừ trang 1.";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };
            const string fixAction = "Vào Insert > Header > Edit Header, bật Different First Page, sau đó ở header từ trang 2 chèn Quick Parts > Document Property > Title để hiển thị tiêu đề \"Integral\".";

            try
            {
                if (!WP03GraderHelpers.HasTitlePageEnabled(studentDocument))
                {
                    WP03GraderHelpers.AddError(result, "Chưa bật Different First Page (w:titlePg), nên chưa loại trừ trang 1 khỏi phần tiêu đề.", fixAction);
                }
                else
                {
                    result.Score += 6m;
                    result.Details.Add("Đã bật Different First Page để trang 1 khác tiêu đề các trang còn lại.");
                }

                if (!WP03GraderHelpers.TryGetDefaultHeaderPart(studentDocument, out var headerXml, out var headerEntry))
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy header mặc định để kiểm tra tiêu đề hiển thị trên các trang.", fixAction);
                    return result;
                }

                result.Score += 3m;
                result.Details.Add($"Đã tìm thấy header mặc định tại part \"{headerEntry}\".");

                var headerText = WP03GraderHelpers.NormalizeText(
                    string.Concat(headerXml.Descendants(WP03GraderHelpers.W + "t").Select(node => node.Value)));
                var coreTitle = WP03GraderHelpers.GetCoreTitle(studentDocument);

                if (string.IsNullOrWhiteSpace(headerText))
                {
                    WP03GraderHelpers.AddError(result, "Header đang rỗng, chưa hiển thị tiêu đề ở các trang từ trang 2 trở đi.", fixAction);
                }
                else if ((!string.IsNullOrWhiteSpace(coreTitle)
                          && headerText.Contains(coreTitle, StringComparison.OrdinalIgnoreCase))
                         || headerText.Contains("Integral", StringComparison.OrdinalIgnoreCase))
                {
                    result.Score += 7m;
                    result.Details.Add($"Header đã hiển thị tiêu đề hợp lệ: \"{headerText}\".");
                }
                else
                {
                    WP03GraderHelpers.AddError(
                        result,
                        $"Header đã có chữ nhưng chưa đúng tiêu đề yêu cầu. Giá trị hiện tại là \"{headerText}\".",
                        fixAction);
                }

                var titleBindingFound = headerXml
                    .Descendants(WP03GraderHelpers.W + "dataBinding")
                    .Any(node =>
                    {
                        var xPath = node.Attribute(WP03GraderHelpers.W + "xpath")?.Value ?? string.Empty;
                        return xPath.Contains("title", StringComparison.OrdinalIgnoreCase);
                    });

                var titleAliasFound = headerXml
                    .Descendants(WP03GraderHelpers.W + "alias")
                    .Any(node => string.Equals(
                        node.Attribute(WP03GraderHelpers.W + "val")?.Value,
                        "Title",
                        StringComparison.OrdinalIgnoreCase));

                if (titleBindingFound || titleAliasFound)
                {
                    result.Score += 4m;
                    result.Details.Add("Header có liên kết Document Property Title, đúng cách chèn tiêu đề động.");
                }
                else
                {
                    WP03GraderHelpers.AddError(result, "Header chưa thể hiện liên kết Document Property Title (alias/binding).", fixAction);
                }
            }
            catch (Exception ex)
            {
                WP03GraderHelpers.AddError(result, $"Lỗi khi chấm Task 1: {ex.Message}.", "Đóng file Word nếu đang mở, kiểm tra file .docx không bị hỏng rồi tải lại để chấm lại Task 1.");
            }

            return result;
        }
    }
}