using System.Xml.Linq;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project01
{
    public class WP01T5Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T5";
        public string TaskName => "Trong phần \"Favorite Dinosaurs\", tại đoạn văn trống ở cuối trang, sử dụng tính năng 3D Models để chèn mô hình \"Triceratops\" từ thư mục 3D Objects. Sau đó định vị mô hình theo kiểu In Line with Text.";
        public decimal MaxScore => 30m;

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
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Favorite Dinosaurs");
                if (headingIndex < 0)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy tiêu đề \"Favorite Dinosaurs\".",
                        "Khôi phục đúng tiêu đề \"Favorite Dinosaurs\" để hệ thống nhận diện vị trí chèn mô hình 3D.");
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
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy đối tượng 3D model trong phần \"Favorite Dinosaurs\".",
                        "Đặt con trỏ ở đoạn trống cuối phần \"Favorite Dinosaurs\", chọn Insert > 3D Models và chèn mô hình Triceratops.");
                    return result;
                }

                result.Score += 8m;
                result.Details.Add("Đã phát hiện đối tượng 3D model trong phần yêu cầu.");

                var modelNode = modelNodes[0];
                var modelRelationshipId = modelNode.Attribute(WP01GraderHelpers.R + "embed")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(modelRelationshipId))
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Đối tượng 3D model chưa có liên kết relationship (r:embed).",
                        "Xóa đối tượng lỗi và chèn lại bằng Insert > 3D Models để Word tạo liên kết mô hình hợp lệ.");
                    return result;
                }

                if (!studentDocument.TryGetDocumentRelationship(modelRelationshipId, out var relationship))
                {
                    WP01GraderHelpers.AddError(
                        result,
                        $"Không tìm thấy relationship id \"{modelRelationshipId}\" cho 3D model.",
                        "Chèn lại mô hình 3D từ Word thay vì copy phần tử từ nguồn khác làm mất liên kết.");
                    return result;
                }

                if (relationship.Type.EndsWith("/model3d", StringComparison.OrdinalIgnoreCase))
                {
                    result.Score += 8m;
                    result.Details.Add("Relationship của đối tượng 3D đúng loại model3d.");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        $"Relationship của 3D model chưa đúng loại. Type hiện tại: \"{relationship.Type}\".",
                        "Đảm bảo đối tượng được chèn là 3D Model Triceratops, không phải ảnh chụp hoặc hình minh họa thường.");
                }

                var modelEntry = WP01GraderHelpers.ResolveWordPartEntry(relationship.Target);
                if (studentDocument.ContainsEntry(modelEntry))
                {
                    result.Score += 4m;
                    result.Details.Add($"Đã tìm thấy tệp mô hình 3D trong package: \"{modelEntry}\".");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        $"Thiếu tệp mô hình 3D trong package. Không tìm thấy \"{modelEntry}\".",
                        "Chèn lại mô hình Triceratops và lưu file .docx để Word nhúng tệp model3d vào tài liệu.");
                }

                var isInline = modelNode.Ancestors(WP01GraderHelpers.WP + "inline").Any();
                if (isInline)
                {
                    result.Score += 4m;
                    result.Details.Add("Mô hình 3D đang ở chế độ In Line with Text (wp:inline).");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Mô hình 3D chưa ở chế độ In Line with Text (wp:inline).",
                        "Chọn mô hình 3D, mở Layout Options/Wrap Text và chọn In Line with Text.");
                }

                var modelParagraph = modelNode.Ancestors(WP01GraderHelpers.W + "p").FirstOrDefault();
                if (modelParagraph == null)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không xác định được đoạn văn chứa mô hình 3D.",
                        "Di chuyển mô hình 3D vào đoạn văn trống trong phần \"Favorite Dinosaurs\" rồi lưu lại.");
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
                    WP01GraderHelpers.AddError(
                        result,
                        "Đoạn văn chứa mô hình 3D không trống hoàn toàn.",
                        "Cắt mô hình 3D và dán vào đoạn trống riêng, không để chung với chữ.");
                }
            }
            catch (Exception ex)
            {
                WP01GraderHelpers.AddError(
                    result,
                    $"Lỗi khi chấm Task 5: {ex.Message}.",
                    "Lưu lại tệp .docx và kiểm tra mô hình Triceratops còn nằm trong phần \"Favorite Dinosaurs\".");
            }

            return result;
        }
    }
}

