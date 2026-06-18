using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project09
{
    public sealed class WP09T2Grader : IWordTaskGrader
    {
        public string TaskId => "W09-T02";
        public string TaskName => "Merge first row cells in Contact us table";
        public decimal MaxScore => 1m;

        public TaskResult Grade(WordGradingContext studentDocument)
        {
            var result = WP09GraderHelpers.CreateResult(TaskId, TaskName, MaxScore);
            var bodyElements = WP09GraderHelpers.GetBodyElements(studentDocument);
            var contactIndex = WP09GraderHelpers.FindParagraphIndexContaining(bodyElements, "Contact us");

            if (contactIndex < 0)
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Không tìm thấy phần Contact us trong tài liệu.",
                    "Khôi phục tiêu đề/phần Contact us, sau đó chọn hàng đầu tiên của bảng trong phần này và dùng Table Layout > Merge Cells.");
                return result;
            }

            var table = WP09GraderHelpers.GetFirstTableAfter(bodyElements, contactIndex);
            if (table == null)
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Không tìm thấy bảng ngay sau phần Contact us.",
                    "Trong phần Contact us, giữ/khôi phục bảng liên hệ rồi chọn tất cả ô của hàng đầu tiên, vào Table Layout > Merge Cells.");
                return result;
            }

            var firstRow = table.Elements(WP09GraderHelpers.W + "tr").FirstOrDefault();
            var firstRowCells = firstRow?.Elements(WP09GraderHelpers.W + "tc").ToList() ?? new List<System.Xml.Linq.XElement>();
            var tableColumnCount = table.Element(WP09GraderHelpers.W + "tblGrid")?
                .Elements(WP09GraderHelpers.W + "gridCol")
                .Count() ?? 0;

            var hasMergedRow = tableColumnCount > 1
                && firstRowCells.Count == 1
                && GetGridSpan(firstRowCells[0]) == tableColumnCount;

            if (!hasMergedRow)
            {
                WP09GraderHelpers.AddError(
                    result,
                    "Hàng đầu tiên của bảng trong phần Contact us chưa được merge thành một ô.",
                    "Trong phần Contact us, chọn tất cả ô của hàng đầu tiên trong bảng, vào Table Layout > Merge Cells, rồi lưu file.");
            }
            else
            {
                result.Details.Add("Hàng đầu tiên của bảng Contact us đã có cấu trúc ô được merge.");
            }

            return result;
        }

        private static int GetGridSpan(System.Xml.Linq.XElement cell)
        {
            var gridSpanValue = cell.Element(WP09GraderHelpers.W + "tcPr")
                ?.Element(WP09GraderHelpers.W + "gridSpan")
                ?.Attribute(WP09GraderHelpers.W + "val")
                ?.Value;

            return int.TryParse(gridSpanValue, out var gridSpan) ? gridSpan : 1;
        }
    }
}

