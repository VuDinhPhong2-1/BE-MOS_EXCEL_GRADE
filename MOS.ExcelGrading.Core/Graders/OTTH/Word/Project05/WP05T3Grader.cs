using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project05
{
    public class WP05T3Grader : IWordTaskGrader
    {
        public string TaskId => "W05-T3";
        public string TaskName => "Trong phần \"Bank Fees\", chuyển đổi văn bản được phân tách bằng dấu tab thành một bảng gồm hai cột. Chấp nhận thiết lập AutoFit mặc định.";
        public decimal MaxScore => 24m;

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
                var bodyElements = WP05GraderHelpers.GetBodyElements(studentDocument);
                var headingIndex = WP05GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Bank Fees");
                if (headingIndex < 0)
                {
                    result.Errors.Add("Không tìm thấy tiêu đề \"Bank Fees\".");
                    result.FixActions.Add("Kiểm tra lại tiêu đề phần \"Bank Fees\"; không đổi tên hoặc xóa tiêu đề này trước khi chuyển văn bản thành bảng.");
                    return result;
                }

                result.Score += 3m;
                result.Details.Add("Đã tìm thấy đúng phần \"Bank Fees\".");

                var table = WP05GraderHelpers.GetFirstTableAfterHeading(bodyElements, headingIndex);
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng trong phần \"Bank Fees\".");
                    result.FixActions.Add("Chọn các dòng dữ liệu được phân tách bằng Tab trong phần \"Bank Fees\", vào Insert > Table > Convert Text to Table và chọn 2 columns.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã có bảng dữ liệu trong phần \"Bank Fees\".");

                var rows = table.Elements(WP05GraderHelpers.W + "tr").ToList();
                if (rows.Count >= 8)
                {
                    result.Score += 6m;
                    result.Details.Add($"Bảng có {rows.Count} dòng dữ liệu, đạt yêu cầu tối thiểu.");
                }
                else
                {
                    result.Errors.Add($"Bảng chưa đủ dữ liệu. Hiện tại chỉ có {rows.Count} dòng.");
                    result.FixActions.Add("Chuyển đổi toàn bộ danh sách phí trong phần \"Bank Fees\" thành bảng, không bỏ sót dòng dữ liệu nào.");
                }

                var invalidRows = rows
                    .Select((row, index) => new
                    {
                        Index = index + 1,
                        CellCount = row.Elements(WP05GraderHelpers.W + "tc").Count()
                    })
                    .Where(item => item.CellCount != 2)
                    .ToList();

                if (invalidRows.Count == 0)
                {
                    result.Score += 5m;
                    result.Details.Add("Tất cả các dòng trong bảng đều có đúng hai cột.");
                }
                else
                {
                    result.Errors.Add(
                        $"Có {invalidRows.Count} dòng không đúng 2 cột. Ví dụ: dòng {invalidRows.First().Index} có {invalidRows.First().CellCount} cột.");
                    result.FixActions.Add("Sửa bảng trong phần \"Bank Fees\" để mỗi dòng có đúng 2 cột: tên loại phí ở cột 1 và số tiền ở cột 2.");
                }

                if (rows.Count > 0)
                {
                    var firstRowCells = rows.First().Elements(WP05GraderHelpers.W + "tc").ToList();
                    var lastRowCells = rows.Last().Elements(WP05GraderHelpers.W + "tc").ToList();

                    var firstRowOk = firstRowCells.Count >= 2
                        && string.Equals(WP05GraderHelpers.GetTableCellText(firstRowCells[0]), "Card Replacement (Loss)", StringComparison.Ordinal)
                        && string.Equals(WP05GraderHelpers.GetTableCellText(firstRowCells[1]), "$ 12", StringComparison.Ordinal);

                    var lastRowOk = lastRowCells.Count >= 2
                        && string.Equals(WP05GraderHelpers.GetTableCellText(lastRowCells[0]), "Bank Transfer: International", StringComparison.Ordinal)
                        && string.Equals(WP05GraderHelpers.GetTableCellText(lastRowCells[1]), "$ 35", StringComparison.Ordinal);

                    if (firstRowOk && lastRowOk)
                    {
                        result.Score += 3m;
                        result.Details.Add("Dữ liệu đầu và cuối bảng đúng thứ tự, đúng chính tả và đúng dấu câu.");
                    }
                    else
                    {
                        result.Errors.Add("Nội dung đầu hoặc cuối bảng chưa đúng yêu cầu về thứ tự, chính tả hoặc dấu câu.");
                        result.FixActions.Add("Kiểm tra lại thứ tự và nội dung bảng: dòng đầu là \"Card Replacement (Loss)\" / \"$ 12\", dòng cuối là \"Bank Transfer: International\" / \"$ 35\".");
                    }
                }

                var sectionParagraphs = WP05GraderHelpers.GetSectionParagraphs(bodyElements, headingIndex, stopAtHeading1: false);
                var tabParagraphCount = sectionParagraphs.Count(WP05GraderHelpers.HasTabCharacter);
                if (tabParagraphCount == 0)
                {
                    result.Score += 3m;
                    result.Details.Add("Không còn đoạn tách cột bằng phím Tab trong phần \"Bank Fees\".");
                }
                else
                {
                    result.Errors.Add($"Vẫn còn {tabParagraphCount} đoạn đang dùng Tab thay vì chuyển thành ô trong bảng.");
                    result.FixActions.Add("Chọn các đoạn còn dùng Tab trong phần \"Bank Fees\" và chuyển chúng vào bảng thay vì để dạng văn bản tab.");
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 3: {ex.Message}.");
                result.FixActions.Add("Kiểm tra tài liệu có mở được trong Word và lưu lại ở định dạng .docx trước khi chấm lại.");
            }

            return result;
        }
    }
}

