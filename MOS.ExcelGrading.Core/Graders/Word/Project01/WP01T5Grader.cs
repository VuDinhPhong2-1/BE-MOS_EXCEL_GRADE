using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project01
{
    public class WP01T5Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T5";
        public string TaskName => "Trong phần \"Favorite Dinosaurs\", tại đoạn văn trống ở cuối trang, sử dụng tính năng 3D Models để chèn mô hình \"Triceratops\" từ thư mục 3D Objects. Sau đó định vị mô hình theo kiểu In Line with Text.";
        public decimal MaxScore => 30m;

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
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Favorite Dinosaurs");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Favorite Dinosaurs\".");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy đúng phần \"Favorite Dinosaurs\".");

                var sectionElements = bodyElements.Skip(headingIndex + 1).TakeWhile(element =>
                {
                    return element.Name != WP01GraderHelpers.W + "p"
                        || !WP01GraderHelpers.IsHeading1Paragraph(element);
                }).ToList();

                var modelNodes = sectionElements
                    .SelectMany(element => element.Descendants(WP01GraderHelpers.AM3D + "model3d"))
                    .ToList();

                if (modelNodes.Count == 0)
                {
                    result.Errors.Add("Không tìm thấy đối tượng 3D model trong phần \"Favorite Dinosaurs\".");
                    return result;
                }

                result.Score += 8m;
                result.Details.Add("Đã phát hiện đối tượng 3D model trong phần yêu cầu.");

                var modelNode = modelNodes[0];
                var modelRelationshipId = modelNode.Attribute(WP01GraderHelpers.R + "embed")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(modelRelationshipId))
                {
                    result.Errors.Add("Đối tượng 3D model chưa có liên kết relationship (r:embed).");
                    return result;
                }

                if (!studentDocument.TryGetDocumentRelationship(modelRelationshipId, out var relationship))
                {
                    result.Errors.Add($"Không tìm thấy relationship id \"{modelRelationshipId}\" cho 3D model.");
                    return result;
                }

                if (relationship.Type.EndsWith("/model3d", StringComparison.OrdinalIgnoreCase))
                {
                    result.Score += 8m;
                    result.Details.Add("Relationship của đối tượng 3D đúng loại model3d.");
                }
                else
                {
                    result.Errors.Add($"Relationship của 3D model chưa đúng loại. Type hiện tại: \"{relationship.Type}\".");
                }

                var modelEntry = WP01GraderHelpers.ResolveWordPartEntry(relationship.Target);
                if (studentDocument.ContainsEntry(modelEntry))
                {
                    result.Score += 4m;
                    result.Details.Add($"Đã tìm thấy tệp mô hình 3D trong package: \"{modelEntry}\".");
                }
                else
                {
                    result.Errors.Add($"Thiếu tệp mô hình 3D trong package. Không tìm thấy \"{modelEntry}\".");
                }

                var isInline = modelNode.Ancestors(WP01GraderHelpers.WP + "inline").Any();
                if (isInline)
                {
                    result.Score += 4m;
                    result.Details.Add("Mô hình 3D đang ở chế độ In Line with Text (wp:inline).");
                }
                else
                {
                    result.Errors.Add("Mô hình 3D chưa ở chế độ In Line with Text (wp:inline).");
                }

                var modelParagraph = modelNode.Ancestors(WP01GraderHelpers.W + "p").FirstOrDefault();
                if (modelParagraph == null)
                {
                    result.Errors.Add("Không xác định được đoạn văn chứa mô hình 3D.");
                    return result;
                }

                var paragraphText = WP01GraderHelpers.GetParagraphText(modelParagraph);
                if (string.IsNullOrWhiteSpace(paragraphText))
                {
                    result.Score += 2m;
                    result.Details.Add("Mô hình 3D được đặt trong đoạn văn trống đúng yêu cầu.");
                }
                else
                {
                    result.Errors.Add("Đoạn văn chứa mô hình 3D không trống hoàn toàn.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 5: {ex.Message}.");
            }

            return result;
        }
    }
}
