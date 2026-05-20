using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    public class WP03T1Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T1";
        public string TaskName => "Hiển thị tiêu đề \"Integral\" trên tất cả các trang của tài liệu, trừ trang 1.";
        public decimal MaxScore => 20m;

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
                if (!WP03GraderHelpers.HasTitlePageEnabled(studentDocument))
                {
                    result.Errors.Add("Chưa bật Different First Page (w:titlePg), nên chưa loại trừ trang 1 khỏi phần tiêu đề.");
                }
                else
                {
                    result.Score += 6m;
                    result.Details.Add("Đã bật Different First Page để trang 1 khác tiêu đề các trang còn lại.");
                }

                if (!WP03GraderHelpers.TryGetDefaultHeaderPart(studentDocument, out var headerXml, out var headerEntry))
                {
                    result.Errors.Add("Không tìm thấy header mặc định để kiểm tra tiêu đề hiển thị trên các trang.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add($"Đã tìm thấy header mặc định tại part \"{headerEntry}\".");

                var headerText = WP03GraderHelpers.NormalizeText(
                    string.Concat(headerXml.Descendants(WP03GraderHelpers.W + "t").Select(node => node.Value)));
                var coreTitle = WP03GraderHelpers.GetCoreTitle(studentDocument);

                if (string.IsNullOrWhiteSpace(headerText))
                {
                    result.Errors.Add("Header đang rỗng, chưa hiển thị tiêu đề ở các trang từ trang 2 trở đi.");
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
                    result.Errors.Add(
                        $"Header đã có chữ nhưng chưa đúng tiêu đề yêu cầu. Giá trị hiện tại là \"{headerText}\".");
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
                    result.Errors.Add("Header chưa thể hiện liên kết Document Property Title (alias/binding).");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 1: {ex.Message}.");
            }

            return result;
        }
    }
}
