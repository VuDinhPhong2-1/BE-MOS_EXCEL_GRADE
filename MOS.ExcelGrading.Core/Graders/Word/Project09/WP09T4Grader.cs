using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project09
{
    public sealed class WP09T4Grader : IWordTaskGrader
    {
        public string TaskId => "W09-T04";
        public string TaskName => "Insert footnote for Event Offers";
        public decimal MaxScore => 1m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP09GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            var footnotesXml = WP09GraderHelpers.GetFootnotesPart(studentDocument);
            var hasExpectedFootnoteText = footnotesXml?.Descendants(WP09GraderHelpers.W + "t")
                .Any(node => WP09GraderHelpers.NormalizeText(node.Value)
                    .Contains("Digital file included", StringComparison.OrdinalIgnoreCase)) == true;

            var eventOffersParagraph = WP09GraderHelpers.GetParagraphs(studentDocument).FirstOrDefault(paragraph =>
                WP09GraderHelpers.GetParagraphText(paragraph).Contains("Event Offers", StringComparison.OrdinalIgnoreCase));

            var hasReferenceNearEventOffers = eventOffersParagraph?.Descendants(WP09GraderHelpers.W + "footnoteReference").Any() == true;

            if (!hasExpectedFootnoteText || !hasReferenceNearEventOffers)
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Chưa tìm thấy footnote đúng nội dung “Digital file included” gắn gần tiêu đề Event Offers.",
                    "Đặt con trỏ cạnh tiêu đề Event Offers, vào References > Insert Footnote, nhập chính xác nội dung footnote: Digital file included.");
            }

            if (result.Errors.Count == 0)
            {
                result.Details.Add("Phát hiện footnote đúng nội dung “Digital file included” được gắn gần tiêu đề Event Offers.");
            }

            return result;
        }
    }
}