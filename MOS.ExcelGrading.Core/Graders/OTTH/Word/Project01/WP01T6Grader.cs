using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project01
{
    public class WP01T6Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T6";
        public string TaskName => "Trong phần \"Dinosaurs in a few points\", áp dụng hiệu ứng nghệ thuật \"Pencil Sketch\" cho bức tranh hóa thạch khủng long.";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var bodyElements = WP01GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Dinosaurs in a few points");
                if (headingIndex < 0)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy tiêu đề \"Dinosaurs in a few points\".",
                        "Khôi phục đúng tiêu đề \"Dinosaurs in a few points\" để hệ thống nhận diện ảnh cần áp dụng hiệu ứng.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Dinosaurs in a few points\".");

                var sectionElements = bodyElements.Skip(headingIndex + 1).TakeWhile(element =>
                {
                    return element.Name != WP01GraderHelpers.W + "p"
                        || !WP01GraderHelpers.IsHeadingParagraph(element);
                }).ToList();

                var drawingsInSection = sectionElements.SelectMany(element => element.Descendants(WP01GraderHelpers.W + "drawing")).ToList();
                if (drawingsInSection.Count == 0)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy hình ảnh nào trong phần \"Dinosaurs in a few points\" để kiểm tra hiệu ứng.",
                        "Khôi phục ảnh hóa thạch khủng long trong phần \"Dinosaurs in a few points\" trước khi áp dụng Artistic Effects.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy hình ảnh trong phần \"Dinosaurs in a few points\".");

                var pencilSketchNodes = sectionElements
                    .SelectMany(element => element.Descendants(WP01GraderHelpers.A14 + "artisticPencilSketch"))
                    .ToList();

                if (pencilSketchNodes.Count == 0)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Chưa phát hiện hiệu ứng nghệ thuật \"Pencil Sketch\" trong phần \"Dinosaurs in a few points\".",
                        "Chọn ảnh hóa thạch khủng long, vào Picture Format > Artistic Effects và chọn Pencil Sketch.");
                    return result;
                }

                result.Score += 9m;
                result.Details.Add("Đã phát hiện hiệu ứng nghệ thuật \"Pencil Sketch\" trong phần yêu cầu.");

                var artisticEffect = pencilSketchNodes[0];
                var imageEmbedId = artisticEffect.Ancestors()
                    .FirstOrDefault(node => node.Name.LocalName == "blip")
                    ?.Attribute(WP01GraderHelpers.R + "embed")
                    ?.Value
                    ?? string.Empty;

                if (string.IsNullOrWhiteSpace(imageEmbedId))
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Hiệu ứng \"Pencil Sketch\" chưa gắn với ảnh hợp lệ (thiếu r:embed).",
                        "Xóa ảnh lỗi, chèn/khôi phục lại ảnh gốc rồi áp dụng lại hiệu ứng Pencil Sketch.");
                    return result;
                }

                if (studentDocument.TryGetDocumentRelationship(imageEmbedId, out var relationship)
                    && relationship.Type.EndsWith("/image", StringComparison.OrdinalIgnoreCase))
                {
                    result.Score += 4m;
                    result.Details.Add("Hiệu ứng \"Pencil Sketch\" đã được gắn với đối tượng ảnh hợp lệ.");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Hiệu ứng \"Pencil Sketch\" không gắn với relationship ảnh hợp lệ.",
                        "Áp dụng Pencil Sketch trực tiếp lên ảnh trong tài liệu, không thay thế bằng đối tượng khác.");
                }
            }
            catch (Exception ex)
            {
                WP01GraderHelpers.AddError(
                    result,
                    $"Lỗi khi chấm Task 6: {ex.Message}.",
                    "Lưu lại tệp .docx và kiểm tra ảnh trong phần \"Dinosaurs in a few points\" còn đọc được.");
            }

            return result;
        }
    }
}

