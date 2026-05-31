using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project01
{
    public class WP01T3Grader : IWordTaskGrader
    {
        public string TaskId => "W01-T3";
        public string TaskName => "Trong phần \"Geological eras\", sắp xếp dữ liệu trong bảng theo cột \"Geologic Period\" tăng dần, sau đó theo cột \"Dinosaur\" cũng theo thứ tự tăng dần.";
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
                var headingIndex = WP01GraderHelpers.FindParagraphIndexByExactText(bodyElements, "Geological eras");
                if (headingIndex < 0)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy tiêu đề \"Geological eras\".",
                        "Khôi phục đúng tiêu đề \"Geological eras\" để hệ thống nhận diện bảng cần sắp xếp.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy đúng phần \"Geological eras\".");

                var table = WP01GraderHelpers.GetFirstTableAfterHeading(bodyElements, headingIndex);
                if (table == null)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không tìm thấy bảng dữ liệu ngay sau phần \"Geological eras\".",
                        "Khôi phục bảng trong phần \"Geological eras\" rồi thực hiện Sort trên chính bảng đó.");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy bảng dữ liệu để kiểm tra sắp xếp.");

                var rows = table.Elements(WP01GraderHelpers.W + "tr").ToList();
                if (rows.Count < 3)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bảng dữ liệu không đủ số dòng để kiểm tra sắp xếp.",
                        "Không xóa dòng dữ liệu trong bảng; khôi phục bảng gốc rồi sắp xếp lại.");
                    return result;
                }

                var headerCells = rows[0].Elements(WP01GraderHelpers.W + "tc").ToList();
                if (headerCells.Count < 2)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Bảng dữ liệu thiếu cột để kiểm tra.",
                        "Khôi phục hai cột Dinosaur và Geologic Period trước khi dùng Sort.");
                    return result;
                }

                var dinosaurHeader = WP01GraderHelpers.GetParagraphText(headerCells[0]);
                var periodHeader = WP01GraderHelpers.GetParagraphText(headerCells[1]);
                if (string.Equals(dinosaurHeader, "Dinosaur", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(periodHeader, "Geologic Period", StringComparison.OrdinalIgnoreCase))
                {
                    result.Score += 4m;
                    result.Details.Add("Tiêu đề cột bảng đúng: \"Dinosaur\" và \"Geologic Period\".");
                }
                else
                {
                    WP01GraderHelpers.AddError(
                        result,
                        $"Tiêu đề cột bảng chưa đúng. Hiện tại là \"{dinosaurHeader}\" và \"{periodHeader}\".",
                        "Giữ nguyên tiêu đề cột là Dinosaur và Geologic Period; chỉ sắp xếp dữ liệu bên dưới hàng tiêu đề.");
                }

                var dataRows = new List<(string Dinosaur, string PeriodText, int PeriodOrder)>();
                for (var i = 1; i < rows.Count; i++)
                {
                    var cells = rows[i].Elements(WP01GraderHelpers.W + "tc").ToList();
                    if (cells.Count < 2)
                    {
                        continue;
                    }

                    var dinosaurValue = WP01GraderHelpers.GetParagraphText(cells[0]);
                    var periodValue = WP01GraderHelpers.GetParagraphText(cells[1]);
                    if (string.IsNullOrWhiteSpace(dinosaurValue) && string.IsNullOrWhiteSpace(periodValue))
                    {
                        continue;
                    }

                    dataRows.Add((dinosaurValue, periodValue, WP01GraderHelpers.ParseGeologicPeriodOrder(periodValue)));
                }

                if (dataRows.Count < 2)
                {
                    WP01GraderHelpers.AddError(
                        result,
                        "Không đủ dữ liệu hàng để kiểm tra thứ tự sắp xếp.",
                        "Khôi phục các dòng dữ liệu trong bảng và không gộp/xóa ô trước khi sắp xếp.");
                    return result;
                }

                var periodSortCorrect = true;
                for (var i = 1; i < dataRows.Count; i++)
                {
                    if (dataRows[i - 1].PeriodOrder > dataRows[i].PeriodOrder)
                    {
                        periodSortCorrect = false;
                        WP01GraderHelpers.AddError(
                            result,
                            $"Thứ tự cột \"Geologic Period\" chưa tăng dần tại dòng dữ liệu {i + 1}: \"{dataRows[i - 1].PeriodText}\" đứng trước \"{dataRows[i].PeriodText}\".",
                            "Chọn bảng > Layout > Sort, Sort by Geologic Period theo Ascending, Then by Dinosaur theo Ascending.");
                        break;
                    }
                }

                if (periodSortCorrect)
                {
                    result.Score += 10m;
                    result.Details.Add("Cột \"Geologic Period\" đã được sắp xếp tăng dần.");
                }

                var secondarySortCorrect = true;
                for (var i = 1; i < dataRows.Count; i++)
                {
                    if (dataRows[i - 1].PeriodOrder != dataRows[i].PeriodOrder)
                    {
                        continue;
                    }

                    var previousDinosaur = WP01GraderHelpers.NormalizeSortText(dataRows[i - 1].Dinosaur);
                    var currentDinosaur = WP01GraderHelpers.NormalizeSortText(dataRows[i].Dinosaur);
                    if (string.Compare(previousDinosaur, currentDinosaur, StringComparison.Ordinal) > 0)
                    {
                        secondarySortCorrect = false;
                        WP01GraderHelpers.AddError(
                            result,
                            $"Sắp xếp phụ theo cột \"Dinosaur\" chưa tăng dần trong nhóm \"{dataRows[i].PeriodText}\": \"{dataRows[i - 1].Dinosaur}\" đứng trước \"{dataRows[i].Dinosaur}\".",
                            "Mở Sort và thêm cấp Then by Dinosaur, Type Text, Ascending.");
                        break;
                    }
                }

                if (secondarySortCorrect)
                {
                    result.Score += 8m;
                    result.Details.Add("Sắp xếp phụ theo cột \"Dinosaur\" trong từng nhóm \"Geologic Period\" là đúng.");
                }
            }
            catch (Exception ex)
            {
                WP01GraderHelpers.AddError(
                    result,
                    $"Lỗi khi chấm Task 3: {ex.Message}.",
                    "Lưu lại tệp .docx và kiểm tra bảng trong phần \"Geological eras\" còn cấu trúc hàng/cột hợp lệ.");
            }

            return result;
        }
    }
}
