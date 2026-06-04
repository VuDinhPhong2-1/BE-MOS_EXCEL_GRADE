using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T1Grader : IWordTaskGrader
    {
        public WP13T1Grader(string taskId = "W13-T01")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Kiá»ƒm tra Accessibility vÃ  thÃªm tiÃªu Ä‘á» cho báº£ng";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var tables = WP13GraderHelpers.GetTables(studentDocument);

            if (tables.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng tÃ¬m tháº¥y báº£ng trong tÃ i liá»‡u Ä‘á»ƒ kiá»ƒm tra lá»—i Accessibility vá» Table Title.",
                    "KhÃ´i phá»¥c Ä‘Ãºng báº£ng trong tÃ i liá»‡u, sau Ä‘Ã³ vÃ o Review > Check Accessibility, chá»n lá»—i liÃªn quan Ä‘áº¿n báº£ng vÃ  Ã¡p dá»¥ng recommended action Ä‘áº§u tiÃªn Ä‘á»ƒ thÃªm Table Title.");
                return result;
            }

            var tablesWithoutTitle = tables
                .Where(table => !WP13GraderHelpers.HasTableTitleOrCaption(table))
                .ToList();

            if (tablesWithoutTitle.Count > 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    $"CÃ³ {tablesWithoutTitle.Count} báº£ng chÆ°a cÃ³ tiÃªu Ä‘á»/caption dÃ¹ng cho Accessibility.",
                    "VÃ o Review > Check Accessibility, má»Ÿ lá»—i Table Title, chá»n recommended action Ä‘áº§u tiÃªn vÃ  nháº­p tiÃªu Ä‘á» phÃ¹ há»£p cho báº£ng.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "Táº¥t cáº£ báº£ng cÃ³ dáº¥u hiá»‡u Table Title/caption hoáº·c metadata mÃ´ táº£ dÃ¹ng cho Accessibility.");
            }

            return result;
        }
    }
}
