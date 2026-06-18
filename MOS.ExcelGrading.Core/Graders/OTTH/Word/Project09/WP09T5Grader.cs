using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project09
{
    public sealed class WP09T5Grader : IWordTaskGrader
    {
        public string TaskId => "W09-T05";
        public string TaskName => "Accept insertion/deletion changes and reject formatting changes";
        public decimal MaxScore => 1m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP09GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);

            if (WP09GraderHelpers.HasUnresolvedTrackedChanges(studentDocument))
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Tài liệu vẫn còn thao tác chèn/xóa hoặc thay đổi định dạng đang được theo dõi chưa xử lý.",
                    "Vào Review > Changes: chọn Accept All Changes để chấp nhận các thao tác chèn thêm và xóa bỏ; sau đó dùng Reject cho mọi formatting changes. Kiểm tra lại để không còn markup trước khi lưu.");
            }

            if (result.Errors.Count == 0)
            {
                result.Details.Add("Không phát hiện thao tác chèn/xóa hoặc thay đổi định dạng đang được theo dõi còn sót lại.");
            }

            return result;
        }
    }
}
