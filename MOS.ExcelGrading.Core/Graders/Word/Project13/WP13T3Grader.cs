using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T3Grader : IWordTaskGrader
    {
        private const int ExpectedWidthTwips = 3170;
        private const int WidthToleranceTwips = 60;

        public WP13T3Grader(string taskId = "W13-T03")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Äáº·t táº¥t cáº£ cá»™t báº£ng Filling Agents rá»™ng 5.59 cm";
        public decimal MaxScore => 25m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var table = WP13GraderHelpers.FindFirstTableAfterHeading(studentDocument, "Filling Agents");

            if (table == null)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng tÃ¬m tháº¥y báº£ng ngay sau tiÃªu Ä‘á»/section â€œFilling Agentsâ€.",
                    "KhÃ´i phá»¥c Ä‘Ãºng báº£ng trong section Filling Agents, chá»n toÃ n bá»™ báº£ng rá»“i Ä‘áº·t Table Layout > Cell Size > Width = 5.59 cm.");
                return result;
            }

            var widths = WP13GraderHelpers.GetTableColumnWidthsTwips(table);
            if (widths.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng Ä‘á»c Ä‘Æ°á»£c Ä‘á»™ rá»™ng cá»™t cá»§a báº£ng Filling Agents trong XML.",
                    "Chá»n toÃ n bá»™ báº£ng Filling Agents, vÃ o Table Layout > Cell Size vÃ  nháº­p Width = 5.59 cm Ä‘á»ƒ Word lÆ°u láº¡i Ä‘á»™ rá»™ng cá»™t.");
                return result;
            }

            var correctCount = widths.Count(width => Math.Abs(width - ExpectedWidthTwips) <= WidthToleranceTwips);
            if (correctCount == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    $"ChÆ°a cÃ³ cá»™t nÃ o cá»§a báº£ng Filling Agents rá»™ng 5.59 cm. GiÃ¡ trá»‹ hiá»‡n táº¡i: {string.Join(", ", widths)} twips.",
                    "Chá»n toÃ n bá»™ báº£ng trong section Filling Agents, khÃ´ng chá»‰ chá»n má»™t cá»™t, rá»“i Ä‘áº·t Table Layout > Cell Size > Width = 5.59 cm.");
            }
            else if (correctCount < widths.Count)
            {
                WP13GraderHelpers.AddError(
                    result,
                    $"Chá»‰ cÃ³ {correctCount}/{widths.Count} cá»™t cá»§a báº£ng Filling Agents rá»™ng 5.59 cm; cÃ³ thá»ƒ báº¡n má»›i chá»‰nh má»™t vÃ i cá»™t.",
                    "BÃ´i Ä‘en toÃ n bá»™ báº£ng Filling Agents rá»“i Ä‘áº·t Width = 5.59 cm Ä‘á»ƒ táº¥t cáº£ cá»™t cÃ³ cÃ¹ng Ä‘á»™ rá»™ng.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, $"Táº¥t cáº£ {widths.Count} cá»™t cá»§a báº£ng Filling Agents cÃ³ Ä‘á»™ rá»™ng xáº¥p xá»‰ 5.59 cm.");
            }

            return result;
        }
    }
}
