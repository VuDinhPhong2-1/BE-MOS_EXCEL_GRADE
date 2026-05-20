using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project01
{
    public class WP01T4Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T4";
        public string TaskName => "Trong phần \"Other points to know about dinosaurs\", thay đổi cấp độ danh sách của mục \"Velociraptor\" thành cấp độ 3.";
        public decimal MaxScore => 15m;

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
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Other points to know about dinosaurs");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Other points to know about dinosaurs\".");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Other points to know about dinosaurs\".");

                var sectionParagraphs = WP01GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: false)
                    .ToList();

                var velociraptorParagraph = WP01GraderHelpers.FindParagraphContainingText(sectionParagraphs, "Velociraptor");
                if (velociraptorParagraph == null)
                {
                    result.Errors.Add("Không tìm thấy mục danh sách \"Velociraptor\" trong phần yêu cầu.");
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
                            result.Errors.Add($"Dòng con thứ {i} bên dưới \"Velociraptor\" có ilvl = \"{childIlvlStr}\", yêu cầu phải là 3 (cấp 4).");
                        }
                    }

                    if (childrenCorrect)
                    {
                        result.Score += 7m;
                        result.Details.Add("Cấp độ danh sách của \"Velociraptor\" đúng cấp 3 (ilvl = 2) và 3 dòng con đúng cấp 4 (ilvl = 3).");
                    }
                    else
                    {
                        result.Errors.Add("\"Velociraptor\" có ilvl = 2 nhưng một hoặc nhiều dòng con (Meaning/Size/Weight) chưa đúng cấp 4 (ilvl = 3).");
                    }
                }
                else
                {
                    result.Errors.Add($"Cấp độ danh sách của \"Velociraptor\" chưa đúng. Giá trị ilvl hiện tại là \"{ilvl}\", yêu cầu là \"2\".");
                }

                if (!string.IsNullOrWhiteSpace(numId))
                {
                    result.Score += 2m;
                    result.Details.Add($"Mục \"Velociraptor\" có liên kết numbering (numId = {numId}).");
                }
                else
                {
                    result.Errors.Add("Mục \"Velociraptor\" chưa có thông tin numbering (numId).");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 4: {ex.Message}.");
            }

            return result;
        }

    }
}
