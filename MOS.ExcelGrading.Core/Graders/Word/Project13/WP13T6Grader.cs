using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T6Grader : IWordTaskGrader
    {
        public WP13T6Grader(string taskId = "W13-T06")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "ChÃ¨n 3D model Blister Packs inline trong Ä‘oáº¡n trá»‘ng cá»§a Description";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var descriptionParagraphs = WP13GraderHelpers.FindParagraphsAfterHeading(studentDocument, "Description");

            if (descriptionParagraphs.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng tÃ¬m tháº¥y section â€œDescriptionâ€ Ä‘á»ƒ kiá»ƒm tra vá»‹ trÃ­ 3D model.",
                    "KhÃ´i phá»¥c section Description, Ä‘áº·t con trá» vÃ o Ä‘oáº¡n vÄƒn trá»‘ng trong section nÃ y, chÃ¨n 3D model Blister Packs vÃ  chá»n In Line with Text.");
                return result;
            }

            var blankParagraphs = descriptionParagraphs
                .Where(paragraph => string.IsNullOrWhiteSpace(WP13GraderHelpers.GetParagraphText(paragraph)))
                .ToList();

            if (blankParagraphs.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng tÃ¬m tháº¥y Ä‘oáº¡n vÄƒn trá»‘ng trong section Description Ä‘á»ƒ chá»©a 3D model.",
                    "Trong section Description, táº¡o/khÃ´i phá»¥c Ä‘oáº¡n vÄƒn trá»‘ng Ä‘Ãºng vá»‹ trÃ­ rá»“i chÃ¨n 3D model Blister Packs vÃ o Ä‘oáº¡n Ä‘Ã³.");
                return result;
            }

            var hasInlineBlisterPacksModel = blankParagraphs
                .Any(paragraph => WP13GraderHelpers.HasInline3DModelInParagraph(paragraph, studentDocument, "Blister Packs"));

            if (!hasInlineBlisterPacksModel)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "ChÆ°a phÃ¡t hiá»‡n 3D model â€œBlister Packsâ€ Ä‘Æ°á»£c chÃ¨n inline trong Ä‘oáº¡n vÄƒn trá»‘ng cá»§a Description.",
                    "Äáº·t con trá» vÃ o Ä‘oáº¡n vÄƒn trá»‘ng trong Description, vÃ o Insert > 3D Models, chá»n Blister Packs, sau Ä‘Ã³ Ä‘áº·t Wrap Text/In Line with Text.");
            }

            if (WP13GraderHelpers.HasAnchored3DModelInScope(descriptionParagraphs))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "PhÃ¡t hiá»‡n 3D model dáº¡ng floating/anchor trong Description; yÃªu cáº§u pháº£i lÃ  In Line with Text.",
                    "Chá»n 3D model Blister Packs, má»Ÿ Layout Options/Wrap Text vÃ  chá»n In Line with Text.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "Äoáº¡n vÄƒn trá»‘ng trong Description cÃ³ 3D model Blister Packs vÃ  Ä‘á»‘i tÆ°á»£ng Ä‘Æ°á»£c Ä‘áº·t In Line with Text.");
            }

            return result;
        }
    }
}
