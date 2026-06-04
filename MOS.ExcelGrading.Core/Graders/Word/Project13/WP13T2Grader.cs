using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project13
{
    public sealed class WP13T2Grader : IWordTaskGrader
    {
        public WP13T2Grader(string taskId = "W13-T02")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Äáº·t trang 3 á»Ÿ hÆ°á»›ng Landscape";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var sectionProperties = WP13GraderHelpers.GetSectionPropertiesInDocumentOrder(studentDocument);

            if (sectionProperties.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "KhÃ´ng tÃ¬m tháº¥y section properties Ä‘á»ƒ xÃ¡c Ä‘á»‹nh hÆ°á»›ng trang.",
                    "Äáº·t con trá» á»Ÿ trang 3, vÃ o Layout > Breaks Ä‘á»ƒ tÃ¡ch section náº¿u cáº§n, sau Ä‘Ã³ chá»n Orientation > Landscape cho riÃªng trang 3.");
                return result;
            }

            var landscapeSections = sectionProperties
                .Select((section, index) => new { Section = section, Index = index })
                .Where(item => WP13GraderHelpers.IsLandscapeSection(item.Section))
                .ToList();

            if (landscapeSections.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "ChÆ°a phÃ¡t hiá»‡n section nÃ o Ä‘Æ°á»£c Ä‘áº·t hÆ°á»›ng Landscape cho trang 3.",
                    "Äáº·t con trá» á»Ÿ trang 3, táº¡o section riÃªng cho trang 3 náº¿u cáº§n, rá»“i vÃ o Layout > Orientation > Landscape.");
            }
            else if (landscapeSections.Count == sectionProperties.Count && sectionProperties.Count > 1)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "ToÃ n bá»™ tÃ i liá»‡u hoáº·c quÃ¡ nhiá»u section Ä‘ang á»Ÿ hÆ°á»›ng Landscape, chÆ°a giá»›i háº¡n riÃªng trang 3.",
                    "Chá»‰ chá»n/tÃ¡ch riÃªng trang 3 rá»“i Ã¡p dá»¥ng Layout > Orientation > Landscape; cÃ¡c trang cÃ²n láº¡i giá»¯ Portrait.");
            }
            else if (landscapeSections.Count > 1)
            {
                WP13GraderHelpers.AddError(
                    result,
                    $"PhÃ¡t hiá»‡n {landscapeSections.Count} section Landscape; yÃªu cáº§u chá»‰ Ã¡p dá»¥ng cho trang 3.",
                    "Kiá»ƒm tra láº¡i section breaks quanh trang 3 vÃ  Ä‘á»•i cÃ¡c section khÃ´ng thuá»™c trang 3 vá» Portrait.");
            }

            if (sectionProperties.Count >= 3)
            {
                var thirdSection = sectionProperties[2];
                if (!WP13GraderHelpers.IsLandscapeSection(thirdSection))
                {
                    WP13GraderHelpers.AddError(
                        result,
                        "Section tÆ°Æ¡ng á»©ng vá»‹ trÃ­ trang/section thá»© 3 chÆ°a á»Ÿ hÆ°á»›ng Landscape.",
                        "Äáº·t con trá» á»Ÿ trang 3 vÃ  Ã¡p dá»¥ng Orientation > Landscape cho section chá»©a trang 3.");
                }
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "TÃ i liá»‡u cÃ³ section Landscape há»£p lá»‡ Ä‘á»ƒ biá»ƒu diá»…n trang 3 theo hÆ°á»›ng ngang.");
            }

            return result;
        }
    }
}
