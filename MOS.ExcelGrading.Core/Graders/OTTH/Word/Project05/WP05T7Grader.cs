using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project05
{
    public class WP05T7Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T7";
        public string TaskName => "Ở cuối phần \"Bank Fees\", liên kết biểu tượng mũi tên với \"Top of the Document\".";
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
                var bodyElements = WP05GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP05GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Bank Fees");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Bank Fees\".");
                    result.FixActions.Add("Kiểm tra lại tiêu đề phần \"Bank Fees\"; không đổi tên hoặc xóa tiêu đề này trước khi tạo liên kết cho biểu tượng.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Bank Fees\".");

                var sectionElements = WP05GraderHelpers.GetSectionElements(bodyElements, headingIndex, stopAtHeading1: false);
                var drawingNodes = sectionElements
                    .SelectMany(element => element.Descendants(WP05GraderHelpers.W + "drawing"))
                    .ToList();

                var linkIds = drawingNodes
                    .Select(WP05GraderHelpers.TryGetHyperlinkIdFromDrawing)
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (linkIds.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy liên kết nào trên biểu tượng trong phần \"Bank Fees\".");
                    result.FixActions.Add("Chọn biểu tượng mũi tên ở cuối phần \"Bank Fees\", vào Insert > Link, chọn Place in This Document > Top of the Document.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add($"Đã tìm thấy {linkIds.Count} liên kết trên đối tượng ở phần \"Bank Fees\".");

                var correctTopLinks = 0;
                var wrongLinks = new List<string>();
                foreach (var linkId in linkIds)
                {
                    if (!studentDocument.TryGetDocumentRelationship(linkId!, out var relation))
                    {
                        wrongLinks.Add($"{linkId}: thiếu relationship.");
                        continue;
                    }

                    var isHyperlink = relation.Type.Contains("/hyperlink", StringComparison.OrdinalIgnoreCase);
                    var isTopTarget = string.Equals(relation.Target, "#_top", StringComparison.Ordinal);

                    if (isHyperlink && isTopTarget)
                    {
                        correctTopLinks++;
                    }
                    else
                    {
                        wrongLinks.Add($"{linkId}: target=\"{relation.Target}\".");
                    }
                }

                if (correctTopLinks > 0)
                {
                    result.Score += 6m;
                    result.Details.Add("Đã liên kết biểu tượng tới đúng đích \"Top of the Document\" (#_top).");
                }
                else
                {
                    result.Errors.Add("Không có liên kết nào trỏ đúng tới \"Top of the Document\" (#_top).");
                    result.FixActions.Add("Sửa hyperlink của biểu tượng để trỏ đến \"Top of the Document\" thay vì URL, email hoặc bookmark khác.");
                }

                if (WP05GraderHelpers.HasBookmark(studentDocument, "_top"))
                {
                    result.Score += 3m;
                    result.Details.Add("Tài liệu có bookmark \"_top\" để liên kết hoạt động chính xác.");
                }
                else
                {
                    result.Errors.Add("Tài liệu chưa có bookmark \"_top\", liên kết lên đầu tài liệu chưa hoàn chỉnh.");
                    result.FixActions.Add("Tạo lại liên kết bằng Insert > Link > Place in This Document > Top of the Document để Word tự tạo/khôi phục đích \"_top\".");
                }

                if (wrongLinks.Count == 0)
                {
                    result.Score += 2m;
                    result.Details.Add("Không phát hiện liên kết sai đích trong biểu tượng của phần này.");
                }
                else
                {
                    result.Errors.Add($"Phát hiện liên kết sai đích: {wrongLinks.First()}");
                    result.FixActions.Add("Kiểm tra các hyperlink trên biểu tượng ở cuối phần \"Bank Fees\" và xóa hoặc sửa mọi liên kết không trỏ đến \"Top of the Document\".");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 7: {ex.Message}.");
                result.FixActions.Add("Kiểm tra tài liệu có mở được trong Word và lưu lại ở định dạng .docx trước khi chấm lại.");
            }

            return result;
        }
    }
}

