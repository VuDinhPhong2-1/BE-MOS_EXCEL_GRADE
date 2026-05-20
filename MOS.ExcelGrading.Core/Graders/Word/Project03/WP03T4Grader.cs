using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    public class WP03T4Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T4";
        public string TaskName => "Trong phần \"Top Sellers\", tiếp tục đánh số danh sách ở đầu cột thứ hai để các mục trong danh sách được đánh số liên tục từ 1 đến 6.";
        public decimal MaxScore => 22m;

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
                var bodyElements = WP03GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Top Sellers");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Top Sellers\".");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Top Sellers\".");

                var table = WP03GraderHelpers.GetFirstTableAfterHeading(bodyElements, headingIndex);
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng danh sách ngay sau phần \"Top Sellers\".");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy bảng danh sách cần kiểm tra.");

                var row = table.Elements(WP03GraderHelpers.W + "tr").FirstOrDefault();
                if (row == null)
                {
                    result.Errors.Add("Bảng không có dòng dữ liệu để kiểm tra đánh số.");
                    return result;
                }

                var cells = row.Elements(WP03GraderHelpers.W + "tc").ToList();
                if (cells.Count < 2)
                {
                    result.Errors.Add("Bảng không đủ 2 cột để kiểm tra đánh số liên tục.");
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
                    result.Errors.Add(
                        $"Số mục đánh số chưa đủ. Cột trái: {leftList.Count}, cột phải: {rightList.Count}.");
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
                    result.Errors.Add($"Có {ilvlInvalidCount} mục không ở cấp ilvl=0.");
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
                    result.Errors.Add("Danh sách trong một cột đang dùng nhiều numId, chưa ổn định để đánh số liên tục.");
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
                    result.Errors.Add(
                        $"Cột thứ hai chưa tiếp tục đánh số từ cột thứ nhất (numId trái={leftId}, phải={rightId}).");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 4: {ex.Message}.");
            }

            return result;
        }
    }
}
