using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project01
{
    public class WP01T1Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T1";
        public string TaskName => "Trong thuộc tính của tệp, thêm \"animals\" vào danh mục (Categories).";
        public decimal MaxScore => 12m;

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
                var categoryValue = studentDocument.CorePropertiesXml?.Root?
                    .Element(WP01GraderHelpers.Cp + "category")
                    ?.Value
                    ?? string.Empty;

                var normalizedCategory = WP01GraderHelpers.NormalizeText(categoryValue);
                if (string.IsNullOrWhiteSpace(normalizedCategory))
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bạn đã không nhập Categories trong thuộc tính tệp.",
                        "Mở File > Info > Properties > Advanced Properties > Summary, nhập animals vào Categories rồi lưu tệp.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add($"Đã tìm thấy Categories: \"{normalizedCategory}\".");

                if (string.Equals(normalizedCategory, "animals", StringComparison.Ordinal))
                {
                    result.Score += 8m;
                    result.Details.Add("Categories đúng chính tả và đúng giá trị \"animals\".");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        $"Bạn đã nhập sai chính tả hoặc dư ký tự. Categories hiện tại là \"{normalizedCategory}\".",
                        "Sửa Categories thành đúng một từ animals, viết thường và không thêm khoảng trắng/ký tự khác.");
                }
            }
            catch (Exception ex)
            {
                WP01GraderHelpers.AddError(
                    result,
                    $"Lỗi khi chấm Task 1: {ex.Message}.",
                    "Đóng Word, mở lại tệp .docx và lưu lại trước khi chấm; nếu lỗi còn lặp lại, kiểm tra tệp có bị hỏng hay không.");
            }

            return result;
        }
    }
}

