using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project03
{
    public class WP03T5Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T5";
        public string TaskName => "Trong phần \"Overview\", áp dụng hiệu ứng \"Soft Round Bevel\" cho đồ họa SmartArt. Hãy chắc chắn chọn toàn bộ SmartArt.";
        public decimal MaxScore => 25m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };
            const string fixAction = "Trong phần Overview, chọn toàn bộ SmartArt, vào SmartArt Design/Format > Shape Effects > Bevel và chọn hiệu ứng Soft Round Bevel để áp dụng cho toàn bộ SmartArt.";

            try
            {
                var bodyElements = WP03GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Overview");
                if (headingIndex < 0)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy tiêu đề \"Overview\".", "Kiểm tra lại tài liệu và đảm bảo vẫn còn tiêu đề \"Overview\" đúng chính tả trước khi áp dụng hiệu ứng cho SmartArt.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Overview\".");

                var sectionElements = WP03GraderHelpers.GetSectionElements(bodyElements, headingIndex, stopAtHeading1: true);
                var smartArtRel = sectionElements
                    .SelectMany(element => element.Descendants(WP03GraderHelpers.Dgm + "relIds"))
                    .FirstOrDefault();

                if (smartArtRel == null)
                {
                    WP03GraderHelpers.AddError(result, "Không tìm thấy SmartArt trong phần Overview để kiểm tra hiệu ứng.", "Khôi phục hoặc chèn đúng SmartArt trong phần Overview, sau đó chọn toàn bộ SmartArt và áp dụng Soft Round Bevel.");
                    return result;
                }

                result.Score += 5m;
                result.Details.Add("Đã tìm thấy đồ họa SmartArt trong phần Overview.");

                var dataModelRelationId = smartArtRel.Attribute(WP03GraderHelpers.R + "dm")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dataModelRelationId))
                {
                    WP03GraderHelpers.AddError(result, "SmartArt không có liên kết dữ liệu r:dm.", "Chọn đúng toàn bộ SmartArt gốc trong phần Overview và áp dụng lại hiệu ứng Soft Round Bevel; tránh chuyển SmartArt thành hình/ảnh rời.");
                    return result;
                }

                if (!WP03GraderHelpers.TryGetRelatedXmlPart(
                        studentDocument,
                        dataModelRelationId,
                        out var smartArtDataXml,
                        out var smartArtDataEntry))
                {
                    WP03GraderHelpers.AddError(
                        result,
                        $"Không mở được part dữ liệu SmartArt từ relationship \"{dataModelRelationId}\".",
                        "Đóng file Word nếu đang mở, kiểm tra SmartArt trong tài liệu không bị hỏng rồi lưu lại file .docx.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add($"Đã mở dữ liệu SmartArt tại part \"{smartArtDataEntry}\".");

                var softRoundCount = smartArtDataXml
                    .Descendants()
                    .SelectMany(node => node.Attributes())
                    .Count(attribute =>
                        string.Equals(attribute.Name.LocalName, "prst", StringComparison.OrdinalIgnoreCase)
                        && string.Equals(attribute.Value, "softRound", StringComparison.OrdinalIgnoreCase));

                if (softRoundCount >= 6)
                {
                    result.Score += 13m;
                    result.Details.Add(
                        $"Đã áp dụng Soft Round Bevel đầy đủ trên SmartArt (phát hiện {softRoundCount} điểm \"softRound\").");
                }
                else if (softRoundCount > 0)
                {
                    result.Score += 6m;
                    WP03GraderHelpers.AddError(
                        result,
                        $"Đã có Soft Round Bevel nhưng chưa đều toàn bộ SmartArt (chỉ phát hiện {softRoundCount} điểm \"softRound\").",
                        fixAction);
                }
                else
                {
                    WP03GraderHelpers.AddError(result, "Chưa phát hiện hiệu ứng Soft Round Bevel trong dữ liệu SmartArt.", fixAction);
                }
            }
            catch (Exception ex)
            {
                WP03GraderHelpers.AddError(result, $"Lỗi khi chấm Task 5: {ex.Message}.", "Đóng file Word nếu đang mở, kiểm tra file .docx không bị hỏng rồi tải lại để chấm lại Task 5.");
            }

            return result;
        }
    }
}

