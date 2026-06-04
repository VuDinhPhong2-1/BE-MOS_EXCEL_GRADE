using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T5Grader : IWordTaskGrader
    {
        public WP13T5Grader(string taskId = "W13-T05")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Äáº·t alt text Process Flow cho SmartArt trong Manufacturing Process";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var manufacturingParagraphs = WP13GraderHelpers.FindParagraphsAfterHeading(studentDocument, "Manufacturing Process");

            if (manufacturingParagraphs.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng tÃ¬m tháº¥y section â€œManufacturing Processâ€ hoáº·c khÃ´ng cÃ³ ná»™i dung Ä‘á»ƒ kiá»ƒm tra SmartArt.",
                    "KhÃ´i phá»¥c section Manufacturing Process, chá»n toÃ n bá»™ SmartArt trong section nÃ y, má»Ÿ Alt Text vÃ  nháº­p Process Flow.");
                return result;
            }

            var sectionScope = manufacturingParagraphs.Cast<System.Xml.Linq.XElement>().ToList();
            if (!sectionScope.Any(paragraph => paragraph.Descendants(WP13GraderHelpers.W + "drawing").Any(WP13GraderHelpers.LooksLikeSmartArt)))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng phÃ¡t hiá»‡n SmartArt trong section â€œManufacturing Processâ€.",
                    "ChÃ¨n/khÃ´i phá»¥c SmartArt Ä‘Ãºng trong section Manufacturing Process, sau Ä‘Ã³ chá»n toÃ n bá»™ SmartArt chá»© khÃ´ng chá»n tá»«ng shape con.");
                return result;
            }

            if (!WP13GraderHelpers.HasDrawingAltTextInScope(sectionScope, "Process Flow", requireSmartArt: true))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "SmartArt trong section â€œManufacturing Processâ€ chÆ°a cÃ³ alt text/title/description lÃ  â€œProcess Flowâ€.",
                    "Chá»n toÃ n bá»™ SmartArt trong Manufacturing Process, má»Ÿ Format/Alt Text vÃ  nháº­p chÃ­nh xÃ¡c Process Flow vÃ o Title hoáº·c Description.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "SmartArt trong Manufacturing Process cÃ³ alt text Process Flow.");
            }

            return result;
        }
    }
}
