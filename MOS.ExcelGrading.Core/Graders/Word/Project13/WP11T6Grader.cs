using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project11
{
    public sealed class WP11T6Grader : IWordTaskGrader
    {
        public string TaskId => "W11-T06";
        public string TaskName => "Chèn 3D model Blister Packs inline trong đoạn trống của Description";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP11GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var descriptionParagraphs = WP11GraderHelpers.FindParagraphsAfterHeading(studentDocument, "Description");

            if (descriptionParagraphs.Count == 0)
            {
                WP11GraderHelpers.AddError(
                    result,
                    "Không tìm thấy section “Description” để kiểm tra vị trí 3D model.",
                    "Khôi phục section Description, đặt con trỏ vào đoạn văn trống trong section này, chèn 3D model Blister Packs và chọn In Line with Text.");
                return result;
            }

            var blankParagraphs = descriptionParagraphs
                .Where(paragraph => string.IsNullOrWhiteSpace(WP11GraderHelpers.GetParagraphText(paragraph)))
                .ToList();

            if (blankParagraphs.Count == 0)
            {
                WP11GraderHelpers.AddError(
                    result,
                    "Không tìm thấy đoạn văn trống trong section Description để chứa 3D model.",
                    "Trong section Description, tạo/khôi phục đoạn văn trống đúng vị trí rồi chèn 3D model Blister Packs vào đoạn đó.");
                return result;
            }

            var hasInlineBlisterPacksModel = blankParagraphs
                .Any(paragraph => WP11GraderHelpers.HasInline3DModelInParagraph(paragraph, studentDocument, "Blister Packs"));

            if (!hasInlineBlisterPacksModel)
            {
                WP11GraderHelpers.AddError(
                    result,
                    "Chưa phát hiện 3D model “Blister Packs” được chèn inline trong đoạn văn trống của Description.",
                    "Đặt con trỏ vào đoạn văn trống trong Description, vào Insert > 3D Models, chọn Blister Packs, sau đó đặt Wrap Text/In Line with Text.");
            }

            if (WP11GraderHelpers.HasAnchored3DModelInScope(descriptionParagraphs))
            {
                WP11GraderHelpers.AddError(
                    result,
                    "Phát hiện 3D model dạng floating/anchor trong Description; yêu cầu phải là In Line with Text.",
                    "Chọn 3D model Blister Packs, mở Layout Options/Wrap Text và chọn In Line with Text.");
            }

            if (result.Errors.Count == 0)
            {
                WP11GraderHelpers.AddDetail(result, "Đoạn văn trống trong Description có 3D model Blister Packs và đối tượng được đặt In Line with Text.");
            }

            return result;
        }
    }
}