using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    public class WP03T5Grader : IWordTaskGrader
    {
        public string TaskId => "W03-T5";
        public string TaskName => "Trong phần \"Overview\", áp dụng hiệu ứng \"Soft Round Bevel\" cho đồ họa SmartArt. Hãy chắc chắn chọn toàn bộ SmartArt.";
        public decimal MaxScore => 25m;

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
                var headingIndex = WP03GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Overview");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Overview\".");
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
                    result.Errors.Add("Không tìm thấy SmartArt trong phần Overview để kiểm tra hiệu ứng.");
                    return result;
                }

                result.Score += 5m;
                result.Details.Add("Đã tìm thấy đồ họa SmartArt trong phần Overview.");

                var dataModelRelationId = smartArtRel.Attribute(WP03GraderHelpers.R + "dm")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dataModelRelationId))
                {
                    result.Errors.Add("SmartArt không có liên kết dữ liệu r:dm.");
                    return result;
                }

                if (!WP03GraderHelpers.TryGetRelatedXmlPart(
                        studentDocument,
                        dataModelRelationId,
                        out var smartArtDataXml,
                        out var smartArtDataEntry))
                {
                    result.Errors.Add(
                        $"Không mở được part dữ liệu SmartArt từ relationship \"{dataModelRelationId}\".");
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
                    result.Errors.Add(
                        $"Đã có Soft Round Bevel nhưng chưa đều toàn bộ SmartArt (chỉ phát hiện {softRoundCount} điểm \"softRound\").");
                }
                else
                {
                    result.Errors.Add("Chưa phát hiện hiệu ứng Soft Round Bevel trong dữ liệu SmartArt.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 5: {ex.Message}.");
            }

            return result;
        }
    }
}
