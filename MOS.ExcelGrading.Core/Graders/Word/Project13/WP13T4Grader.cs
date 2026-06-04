using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T4Grader : IWordTaskGrader
    {
        public WP13T4Grader(string taskId = "W13-T04")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "ThÃªm placeholder citation Fabrication1 cuá»‘i Ä‘oáº¡n vÄƒn thá»© hai trong Description";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var descriptionParagraphs = WP13GraderHelpers.FindParagraphsAfterHeading(studentDocument, "Description")
                .Where(paragraph => !string.IsNullOrWhiteSpace(WP13GraderHelpers.GetParagraphText(paragraph)))
                .ToList();

            if (descriptionParagraphs.Count < 2)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng tÃ¬m tháº¥y Ä‘oáº¡n vÄƒn thá»© hai dÆ°á»›i tiÃªu Ä‘á» â€œDescriptionâ€.",
                    "KhÃ´i phá»¥c ná»™i dung section Description, Ä‘áº·t con trá» á»Ÿ cuá»‘i Ä‘oáº¡n vÄƒn thá»© hai rá»“i vÃ o References > Insert Citation > Add New Placeholder vÃ  nháº­p Fabrication1.");
                return result;
            }

            var secondParagraph = descriptionParagraphs[1];
            if (!WP13GraderHelpers.HasCitationPlaceholderAtEnd(secondParagraph, "Fabrication1"))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "ChÆ°a phÃ¡t hiá»‡n placeholder citation â€œFabrication1â€ á»Ÿ cuá»‘i Ä‘oáº¡n vÄƒn thá»© hai dÆ°á»›i Description.",
                    "Äáº·t con trá» á»Ÿ cuá»‘i Ä‘oáº¡n vÄƒn thá»© hai trong section Description, vÃ o References > Insert Citation > Add New Placeholder, nháº­p Fabrication1 rá»“i lÆ°u tÃ i liá»‡u.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "Äoáº¡n vÄƒn thá»© hai trong Description cÃ³ placeholder citation Fabrication1 á»Ÿ cuá»‘i Ä‘oáº¡n.");
            }

            return result;
        }
    }
}
