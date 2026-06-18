using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project13
{
    public sealed class WP13T6Grader : IWordTaskGrader
    {
        public WP13T6Grader(string taskId = "W13-T06")
        {
            TaskId = taskId;
        }

        public string TaskId { get; }
        public string TaskName => "Chèn 3D model Blister Packs inline trong đoạn trống của Description";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP13GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var descriptionParagraphs = WP13GraderHelpers.FindParagraphsAfterHeading(studentDocument, "Description");

            if (descriptionParagraphs.Count == 0)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Không tìm thấy section “Description” để kiểm tra vị trí 3D model.",
                    "Khôi phục section Description, đặt con trỏ vào đoạn văn trống trong section này, chèn 3D model Blister Packs và chọn In Line with Text.");
                return result;
            }

            var blankParagraphs = descriptionParagraphs
                .Where(paragraph => string.IsNullOrWhiteSpace(WP13GraderHelpers.GetParagraphText(paragraph)))
                .ToList();

            var hasInlineModelInBlankParagraph = blankParagraphs
                .Any(paragraph => WP13GraderHelpers.HasInline3DModelInParagraph(paragraph, studentDocument));

            var hasAnyInline3DModelInDescription = descriptionParagraphs
                .Any(paragraph => WP13GraderHelpers.HasInline3DModelInParagraph(paragraph, studentDocument));

            if (blankParagraphs.Count == 0 && !hasAnyInline3DModelInDescription)
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Không tìm thấy đoạn văn trống trong section Description để chứa 3D model.",
                    "Trong section Description, tạo/khôi phục đoạn văn trống đúng vị trí rồi chèn 3D model Blister Packs vào đoạn đó.");
                return result;
            }

            if (!hasInlineModelInBlankParagraph)
            {
                WP13GraderHelpers.AddError(
                    result,
                    hasAnyInline3DModelInDescription
                        ? "Phát hiện 3D model trong section Description nhưng không nằm trong đoạn văn trống yêu cầu."
                        : "Chưa phát hiện 3D model được chèn inline trong đoạn văn trống của Description.",
                    "Đặt con trỏ vào đoạn văn trống trong Description, vào Insert > 3D Models, chọn Blister Packs, sau đó đặt Wrap Text/In Line with Text.");
            }

            if (!hasInlineModelInBlankParagraph && WP13GraderHelpers.HasAnchored3DModelInScope(descriptionParagraphs))
            {
                WP13GraderHelpers.AddError(
                    result,
                    "Phát hiện 3D model dạng floating/anchor trong Description; yêu cầu phải là In Line with Text.",
                    "Chọn 3D model Blister Packs, mở Layout Options/Wrap Text và chọn In Line with Text.");
            }

            if (result.Errors.Count == 0)
            {
                WP13GraderHelpers.AddDetail(result, "Đoạn trống trong Description có 3D model và đối tượng được đặt In Line with Text.");
            }

            return result;
        }
    }
}

