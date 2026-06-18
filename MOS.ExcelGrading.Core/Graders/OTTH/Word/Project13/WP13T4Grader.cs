using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project13
{
    public sealed class WP13T4Grader : IWordTaskGrader
    {
        public WP13T4Grader(string taskId = "W13-T04")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Thêm placeholder citation Fabrication1 cuối đoạn văn thứ hai trong Description";
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
                    "Không tìm thấy đoạn văn thứ hai dưới tiêu đề “Description”.",
                    "Khôi phục nội dung section Description, đặt con trỏ ở cuối đoạn văn thứ hai rồi vào References > Insert Citation > Add New Placeholder và nhập Fabrication1.");
                return result;
            }

            var secondParagraph = descriptionParagraphs[1];
            if (!WP13GraderHelpers.HasCitationPlaceholderAtEnd(secondParagraph, "Fabrication1"))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Chưa phát hiện placeholder citation “Fabrication1” ở cuối đoạn văn thứ hai dưới Description.",
                    "Đặt con trỏ ở cuối đoạn văn thứ hai trong section Description, vào References > Insert Citation > Add New Placeholder, nhập Fabrication1 rồi lưu tài liệu.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "Đoạn văn thứ hai trong Description có placeholder citation Fabrication1 ở cuối đoạn.");
            }

            return result;
        }
    }
}

