using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project01
{
    public class WP01T6Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T6";
        public string TaskName => "Trong phần \"Dinosaurs in a few points\", áp dụng hiệu ứng nghệ thuật \"Pencil Sketch\" cho bức tranh hóa thạch khủng long.";
        public decimal MaxScore => 20m;

        public TaskResult Grade(WordGradingContext studentDocument, WordGradingContext? answerDocument = null)
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
                    result.Errors.Add("Không tìm thấy tiêu đề \"Dinosaurs in a few points\".");
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
                    result.Errors.Add("Không tìm thấy hình ảnh nào trong phần \"Dinosaurs in a few points\" để kiểm tra hiệu ứng.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy hình ảnh trong phần \"Dinosaurs in a few points\".");

                var pencilSketchNodes = sectionElements
                    .SelectMany(element => element.Descendants(WP01GraderHelpers.A14 + "artisticPencilSketch"))
                    .ToList();

                if (pencilSketchNodes.Count == 0)
                {
                    result.Errors.Add("Chưa phát hiện hiệu ứng nghệ thuật \"Pencil Sketch\" trong phần \"Dinosaurs in a few points\".");
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
                    result.Errors.Add("Hiệu ứng \"Pencil Sketch\" chưa gắn với ảnh hợp lệ (thiếu r:embed).");
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
                    result.Errors.Add("Hiệu ứng \"Pencil Sketch\" không gắn với relationship ảnh hợp lệ.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 6: {ex.Message}.");
            }

            return result;
        }
    }
}
