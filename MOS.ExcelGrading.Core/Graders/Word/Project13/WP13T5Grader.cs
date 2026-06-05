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
        public string TaskName => "Đặt alt text Process Flow cho SmartArt trong Manufacturing Process";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var manufacturingParagraphs = WP13GraderHelpers.FindParagraphsAfterHeading(studentDocument, "Manufacturing Process");

            if (manufacturingParagraphs.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Không tìm thấy section \"Manufacturing Process\" hoặc không có nội dung để kiểm tra SmartArt.",
                    "Khôi phục section Manufacturing Process, chọn toàn bộ SmartArt trong section này, mở Alt Text và nhập Process Flow.");
                return result;
            }

            var sectionScope = manufacturingParagraphs.Cast<System.Xml.Linq.XElement>().ToList();
            if (!sectionScope.Any(paragraph => paragraph.Descendants(WP13GraderHelpers.W + "drawing").Any(WP13GraderHelpers.LooksLikeSmartArt)))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Không phát hiện SmartArt trong section \"Manufacturing Process\".",
                    "Chèn/khôi phục SmartArt đúng trong section Manufacturing Process, sau đó chọn toàn bộ SmartArt chứ không chọn từng shape con.");
                return result;
            }

            if (!WP13GraderHelpers.HasDrawingAltTextInScope(sectionScope, "Process Flow", requireSmartArt: true))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "SmartArt trong section \"Manufacturing Process\" chưa có alt text/title/description là \"Process Flow\".",
                    "Chọn toàn bộ SmartArt trong Manufacturing Process, mở Format/Alt Text và nhập chính xác Process Flow vào Title hoặc Description.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "SmartArt trong Manufacturing Process có alt text Process Flow.");
            }

            return result;
        }
    }
}
