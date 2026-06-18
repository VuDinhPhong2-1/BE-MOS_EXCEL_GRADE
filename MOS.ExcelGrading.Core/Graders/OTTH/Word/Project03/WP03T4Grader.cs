using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project03
{
    public class WP03T4Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T4";
        public string TaskName => "Trong phần \"Top Sellers\", tiếp tục đánh số danh sách ở đầu cột thứ hai để các mục trong danh sách được đánh số liên tục từ 1 đến 6.";
        public decimal MaxScore => 22m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };
            const string fixAction = "Trong phần Top Sellers, chọn mục đầu tiên ở cột thứ hai của danh sách, nhấp chuột phải vào số thứ tự và chọn Continue Numbering để danh sách tiếp tục đánh số liên tục từ 1 đến 6.";

            try
            {
                var bodyElements = WP03GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Top Sellers");
                if (headingIndex < 0)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy tiêu đề \"Top Sellers\".", "Kiểm tra lại tài liệu và đảm bảo vẫn còn tiêu đề \"Top Sellers\" đúng chính tả trước khi chỉnh danh sách đánh số.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Top Sellers\".");

                var table = WP03GraderHelpers.GetFirstTableAfterHeading(bodyElements, headingIndex);
                if (table == null)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy bảng danh sách ngay sau phần \"Top Sellers\".", "Khôi phục bảng danh sách trong phần Top Sellers, sau đó áp dụng Continue Numbering cho cột thứ hai.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy bảng danh sách cần kiểm tra.");

                var row = table.Elements(WP03GraderHelpers.W + "tr").FirstOrDefault();
                if (row == null)
                {
                    WP03GraderHelpers.AddError(result, "Bảng không có dòng dữ liệu để kiểm tra đánh số.", "Kiểm tra lại bảng trong phần Top Sellers để đảm bảo bảng còn các dòng/cột chứa danh sách đánh số.");
                    return result;
                }

                var cells = row.Elements(WP03GraderHelpers.W + "tc").ToList();
                if (cells.Count < 2)
                {
                    WP03GraderHelpers.AddError(result, "Bảng không đủ 2 cột để kiểm tra đánh số liên tục.", "Khôi phục bảng Top Sellers có đủ hai cột danh sách, rồi dùng Continue Numbering cho danh sách ở cột thứ hai.");
                    return result;
                }

                var leftList = cells[0].Elements(WP03GraderHelpers.W + "p")
                    .Where(paragraph => !string.IsNullOrWhiteSpace(WP03GraderHelpers.GetNumId(paragraph)))
                    .ToList();
                var rightList = cells[1].Elements(WP03GraderHelpers.W + "p")
                    .Where(paragraph => !string.IsNullOrWhiteSpace(WP03GraderHelpers.GetNumId(paragraph)))
                    .ToList();

                if (leftList.Count >= 3 && rightList.Count >= 3)
                {
                    result.Score += 4m;
                    result.Details.Add($"Danh sách ở hai cột đã nhận diện được {leftList.Count + rightList.Count} mục đánh số.");
                }
                else
                {
                    WP03GraderHelpers.AddError(
                        result,
                        $"Số mục đánh số chưa đủ. Cột trái: {leftList.Count}, cột phải: {rightList.Count}.",
                        "Đảm bảo mỗi cột trong bảng Top Sellers có đủ 3 mục được định dạng Numbering, tổng cộng 6 mục trong danh sách.");
                    return result;
                }

                var ilvlInvalidCount = leftList.Concat(rightList)
                    .Count(paragraph => !string.Equals(WP03GraderHelpers.GetIlvl(paragraph), "0", StringComparison.Ordinal));
                if (ilvlInvalidCount == 0)
                {
                    result.Score += 3m;
                    result.Details.Add("Tất cả mục danh sách đều ở cấp ilvl=0.");
                }
                else
                {
                    WP03GraderHelpers.AddError(result, $"Có {ilvlInvalidCount} mục không ở cấp ilvl=0.", "Chọn các mục trong danh sách Top Sellers và áp dụng cùng một cấp đánh số chính (level 1), không dùng thụt cấp phụ.");
                }

                var leftNumIds = leftList.Select(WP03GraderHelpers.GetNumId).Distinct(StringComparer.Ordinal).ToList();
                var rightNumIds = rightList.Select(WP03GraderHelpers.GetNumId).Distinct(StringComparer.Ordinal).ToList();

                if (leftNumIds.Count == 1 && rightNumIds.Count == 1)
                {
                    result.Score += 2m;
                    result.Details.Add(
                        $"Mỗi cột đang dùng một numId ổn định. Cột trái={leftNumIds[0]}, cột phải={rightNumIds[0]}.");
                }
                else
                {
                    WP03GraderHelpers.AddError(result, "Danh sách trong một cột đang dùng nhiều numId, chưa ổn định để đánh số liên tục.", fixAction);
                }

                if (leftNumIds.Count == 1 && rightNumIds.Count == 1
                    && string.Equals(leftNumIds[0], rightNumIds[0], StringComparison.Ordinal))
                {
                    result.Score += 6m;
                    result.Details.Add("Cột thứ hai đã tiếp tục đúng danh sách đánh số từ cột thứ nhất.");
                }
                else
                {
                    var leftId = leftNumIds.Count == 1 ? leftNumIds[0] : "không xác định";
                    var rightId = rightNumIds.Count == 1 ? rightNumIds[0] : "không xác định";
                    WP03GraderHelpers.AddError(
                        result,
                        $"Cột thứ hai chưa tiếp tục đánh số từ cột thứ nhất (numId trái={leftId}, phải={rightId}).",
                        fixAction);
                }
            }
            catch (Exception ex)
            {
                WP03GraderHelpers.AddError(result, $"Lỗi khi chấm Task 4: {ex.Message}.", "Đóng file Word nếu đang mở, kiểm tra file .docx không bị hỏng rồi tải lại để chấm lại Task 4.");
            }

            return result;
        }
    }
}

