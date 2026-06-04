using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project09
{
    public sealed class WP09T3Grader : IWordTaskGrader
    {
        private static readonly string[] ExpectedItems =
        {
            "Corporate events",
            "School events",
            "Sporting events",
            "Weddings",
            "Religious ceremonies"
        };

        public string TaskId => "W09-T03";
        public string TaskName => "Convert five event paragraphs to default bullets";
        public decimal MaxScore => 1m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP09GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var paragraphs = WP09GraderHelpers.GetParagraphs(studentDocument);

            var contactUsParagraph = paragraphs.FirstOrDefault(p =>
                WP09GraderHelpers.GetParagraphText(p).StartsWith("Contact Us", StringComparison.OrdinalIgnoreCase));

            if (contactUsParagraph != null && WP09GraderHelpers.HasParagraphNumbering(contactUsParagraph))
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Tiêu đề Contact Us đang bị áp dụng bullet mặc định.",
                    "Chỉ áp dụng Home > Bullets với kiểu dấu đầu dòng mặc định cho 5 đoạn Corporate events, School events, Sporting events, Weddings và Religious ceremonies; không áp dụng bullet cho tiêu đề Contact Us.");
            }

            var missingOrUnbulleted = new List<string>();
            foreach (var expectedItem in ExpectedItems)
            {
                var paragraph = paragraphs.FirstOrDefault(p =>
                    WP09GraderHelpers.GetParagraphText(p).StartsWith(expectedItem, StringComparison.OrdinalIgnoreCase));

                if (paragraph == null || !WP09GraderHelpers.HasParagraphNumbering(paragraph))
                {
                    missingOrUnbulleted.Add(expectedItem);
                }
            }

            if (missingOrUnbulleted.Count > 0)
            {
                WP09GraderHelpers.AddError(
                    result,
                    $"Các đoạn sau chưa có bullet mặc định hoặc không tìm thấy: {string.Join(", ", missingOrUnbulleted)}.",
                    "Trong phần Preserve your greatest memories!, chọn 5 đoạn Corporate events, School events, Sporting events, Weddings và Religious ceremonies, sau đó chọn Home > Bullets với kiểu dấu đầu dòng mặc định.");
            }
            else
            {
                result.Details.Add("Cả 5 đoạn sự kiện đều có numbering/bullet trong OpenXML và tiêu đề Contact Us không bị áp dụng bullet.");
            }

            return result;
        }
    }
}