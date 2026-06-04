using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project01
{
    public class WP01T4Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T4";
        public string TaskName => "Trong phần \"Other points to know about dinosaurs\", thay đổi cấp độ danh sách của mục \"Velociraptor\" thành cấp độ 3.";
        public decimal MaxScore => 15m;

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
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Other points to know about dinosaurs");
                if (headingIndex < 0)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy tiêu đề \"Other points to know about dinosaurs\".",
                        "Khôi phục đúng tiêu đề \"Other points to know about dinosaurs\" để hệ thống nhận diện danh sách.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Other points to know about dinosaurs\".");

                var sectionParagraphs = WP01GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: false)
                    .ToList();

                var velociraptorParagraph = WP01GraderHelpers.FindParagraphContainingText(sectionParagraphs, "Velociraptor");
                if (velociraptorParagraph == null)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy mục danh sách \"Velociraptor\" trong phần yêu cầu.",
                        "Khôi phục mục danh sách \"Velociraptor\" và các dòng con Meaning/Size/Weight.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy mục danh sách \"Velociraptor\".");

                var ilvl = velociraptorParagraph.Element(WP01GraderHelpers.W + "pPr")
                    ?.Element(WP01GraderHelpers.W + "numPr")
                    ?.Element(WP01GraderHelpers.W + "ilvl")
                    ?.Attribute(WP01GraderHelpers.W + "val")
                    ?.Value
                    ?? string.Empty;

                var numId = velociraptorParagraph.Element(WP01GraderHelpers.W + "pPr")
                    ?.Element(WP01GraderHelpers.W + "numPr")
                    ?.Element(WP01GraderHelpers.W + "numId")
                    ?.Attribute(WP01GraderHelpers.W + "val")
                    ?.Value
                    ?? string.Empty;

                if (string.Equals(ilvl, "2", StringComparison.Ordinal))
                {
                    var velociraptorIndex = sectionParagraphs.IndexOf(velociraptorParagraph);
                    bool childrenCorrect = true;

                    for (int i = 1; i <= 3 && (velociraptorIndex + i) < sectionParagraphs.Count; i++)
                    {
                        var childPara = sectionParagraphs[velociraptorIndex + i];
                        var childIlvlStr = childPara.Element(WP01GraderHelpers.W + "pPr")
                            ?.Element(WP01GraderHelpers.W + "numPr")
                            ?.Element(WP01GraderHelpers.W + "ilvl")
                            ?.Attribute(WP01GraderHelpers.W + "val")
                            ?.Value
                            ?? string.Empty;

                        // ✅ Phải đúng bằng ilvl = 3 (cấp 4)
                        if (!int.TryParse(childIlvlStr, out int childIlvl) || childIlvl != 3)
                        {
                            childrenCorrect = false;
                            WP01GraderHelpers.AddError(
                                result,
                                $"Dòng con thứ {i} bên dưới \"Velociraptor\" có ilvl = \"{childIlvlStr}\", yêu cầu phải là 3 (cấp 4).",
                                "Sau khi đưa \"Velociraptor\" về cấp 3, giữ các dòng Meaning/Size/Weight bên dưới ở cấp 4.");
                        }
                    }

                    if (childrenCorrect)
                    {
                        result.Score += 7m;
                        result.Details.Add("Cấp độ danh sách của \"Velociraptor\" đúng cấp 3 (ilvl = 2) và 3 dòng con đúng cấp 4 (ilvl = 3).");
                    }
                    else
                    {
                        WP01GraderHelpers.AddError(
                            result,
                            "\"Velociraptor\" có ilvl = 2 nhưng một hoặc nhiều dòng con (Meaning/Size/Weight) chưa đúng cấp 4 (ilvl = 3).",
                            "Dùng Increase Indent/Change List Level để Meaning, Size và Weight nằm dưới \"Velociraptor\" một cấp.");
                    }
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        $"Cấp độ danh sách của \"Velociraptor\" chưa đúng. Giá trị ilvl hiện tại là \"{ilvl}\", yêu cầu là \"2\".",
                        "Chọn dòng \"Velociraptor\", vào Multilevel List/Change List Level và đặt thành cấp độ 3.");
                }

                if (!string.IsNullOrWhiteSpace(numId))
                {
                    result.Score += 2m;
                    result.Details.Add($"Mục \"Velociraptor\" có liên kết numbering (numId = {numId}).");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Mục \"Velociraptor\" chưa có thông tin numbering (numId).",
                        "Áp dụng lại danh sách đa cấp cho mục \"Velociraptor\" thay vì gõ thụt lề thủ công.");
                }
            }
            catch (Exception ex)
            {
                WP01GraderHelpers.AddError(
                    result,
                    $"Lỗi khi chấm Task 4: {ex.Message}.",
                    "Lưu lại tệp .docx và kiểm tra danh sách đa cấp trong phần \"Other points to know about dinosaurs\".");
            }

            return result;
        }

    }
}
