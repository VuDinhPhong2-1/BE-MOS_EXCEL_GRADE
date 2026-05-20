using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project01
{
    public class WP01T1Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T1";
        public string TaskName => "Trong thuộc tính của tệp, thêm \"animals\" vào danh mục (Categories).";
        public decimal MaxScore => 12m;

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
                var categoryValue = studentDocument.CorePropertiesXml?.Root?
                    .Element(WP01GraderHelpers.Cp + "category")
                    ?.Value
                    ?? string.Empty;

                var normalizedCategory = WP01GraderHelpers.NormalizeText(categoryValue);
                if (string.IsNullOrWhiteSpace(normalizedCategory))
                {
                    result.Errors.Add("Chưa thiết lập Categories trong thuộc tính tệp.");
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
                    result.Errors.Add($"Categories chưa đúng. Giá trị hiện tại là \"{normalizedCategory}\", yêu cầu chính xác là \"animals\".");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 1: {ex.Message}.");
            }

            return result;
        }
    }
}
