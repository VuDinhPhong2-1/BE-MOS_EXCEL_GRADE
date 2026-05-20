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
                    result.Errors.Add("Không tìm thấy tiêu đề \"Geological eras\".");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy đúng phần \"Geological eras\".");

                var table = WP01GraderHelpers.GetFirstTableAfterHeading(bodyElements, headingIndex);
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu ngay sau phần \"Geological eras\".");
                    return result;
                }

                result.Score += 4m;
                result.Details.Add("Đã tìm thấy bảng dữ liệu để kiểm tra sắp xếp.");

                var rows = table.Elements(WP01GraderHelpers.W + "tr").ToList();
                if (rows.Count < 3)
                {
                    result.Errors.Add("Bảng dữ liệu không đủ số dòng để kiểm tra sắp xếp.");
                    return result;
                }

                var headerCells = rows[0].Elements(WP01GraderHelpers.W + "tc").ToList();
                if (headerCells.Count < 2)
                {
                    result.Errors.Add("Bảng dữ liệu thiếu cột để kiểm tra.");
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
                    result.Errors.Add($"Tiêu đề cột bảng chưa đúng. Hiện tại là \"{dinosaurHeader}\" và \"{periodHeader}\".");
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
                    result.Errors.Add("Không đủ dữ liệu hàng để kiểm tra thứ tự sắp xếp.");
                    return result;
                }

                var periodSortCorrect = true;
                for (var i = 1; i < dataRows.Count; i++)
                {
                    if (dataRows[i - 1].PeriodOrder > dataRows[i].PeriodOrder)
                    {
                        periodSortCorrect = false;
                        result.Errors.Add(
                            $"Thứ tự cột \"Geologic Period\" chưa tăng dần tại dòng dữ liệu {i + 1}: \"{dataRows[i - 1].PeriodText}\" đứng trước \"{dataRows[i].PeriodText}\".");
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
                        result.Errors.Add(
                            $"Sắp xếp phụ theo cột \"Dinosaur\" chưa tăng dần trong nhóm \"{dataRows[i].PeriodText}\": \"{dataRows[i - 1].Dinosaur}\" đứng trước \"{dataRows[i].Dinosaur}\".");
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
                result.Errors.Add($"Lỗi khi chấm Task 3: {ex.Message}.");
            }

            return result;
        }
    }
}
