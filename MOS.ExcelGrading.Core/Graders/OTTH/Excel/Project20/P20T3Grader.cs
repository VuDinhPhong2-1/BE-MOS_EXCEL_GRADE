using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project20
{
    public class P20T3Grader : ITaskGrader
    {
        public string TaskId => "P20-T3";
        public string TaskName => "Trong trang tính “New York City”, sắp xếp dữ liệu trong bảng theo nhiều cấp độ: trước tiên theo Country or region (từ A đến Z), sau đó theo City (từ A đến Z).";
        public decimal MaxScore => 24m;

        public TaskResult Grade(ExcelWorksheet studentSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var worksheet = P20GraderHelpers.GetSheet(studentSheet.Workbook, "New York City");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'New York City'.");
                    return result;
                }

                var table = P20GraderHelpers.FindTable(
                    worksheet,
                    "Table14",
                    "Country or region",
                    "City",
                    "Airport code",
                    "Air Miles");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng Table14 trên sheet 'New York City'.");
                    return result;
                }

                if (!P20GraderHelpers.TryGetColumnIndex(table, "Country or region", out var countryCol)
                    || !P20GraderHelpers.TryGetColumnIndex(table, "City", out var cityCol))
                {
                    result.Errors.Add("Không xác định được cột 'Country or region' hoặc cột 'City' trong bảng.");
                    return result;
                }

                decimal score = 0m;

                if (P20GraderHelpers.TryGetTableSortState(table, out var sortRef, out var conditionRefs))
                {
                    var hasSortRef = P20GraderHelpers.IsRangeMatch(sortRef, "A5:D21");
                    var hasTwoLevels = conditionRefs.Count == 2
                                       && P20GraderHelpers.IsRangeMatch(conditionRefs[0], "A5:A21")
                                       && P20GraderHelpers.IsRangeMatch(conditionRefs[1], "B5:B21");

                    if (hasSortRef && hasTwoLevels)
                    {
                        score += 8m;
                        result.Details.Add("Bảng đã lưu đúng cấu hình sắp xếp 2 cấp: Country or region rồi City (A→Z).");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"SortState của bảng chưa đúng. ref='{sortRef}', conditions=[{string.Join(", ", conditionRefs)}].");
                    }
                }
                else
                {
                    result.Errors.Add("Bảng chưa có thông tin SortState cho thao tác sắp xếp nhiều cấp.");
                }

                var firstDataRow = table.Address.Start.Row + 1;
                var lastDataRow = table.Address.End.Row;
                var totalRows = Math.Max(0, lastDataRow - firstDataRow + 1);
                if (totalRows <= 0)
                {
                    result.Errors.Add("Bảng New York City không có dữ liệu để chấm.");
                    result.Score = score;
                    return result;
                }

                var sortedPairs = 0;
                var totalPairs = Math.Max(0, totalRows - 1);
                var firstViolation = string.Empty;

                for (var row = firstDataRow; row < lastDataRow; row++)
                {
                    var countryCurrent = worksheet.Cells[row, countryCol].Text ?? string.Empty;
                    var cityCurrent = worksheet.Cells[row, cityCol].Text ?? string.Empty;
                    var countryNext = worksheet.Cells[row + 1, countryCol].Text ?? string.Empty;
                    var cityNext = worksheet.Cells[row + 1, cityCol].Text ?? string.Empty;

                    var countryCompare = P20GraderHelpers.CompareSortText(countryCurrent, countryNext);
                    var cityCompare = P20GraderHelpers.CompareSortText(cityCurrent, cityNext);

                    var pairIsSorted = countryCompare < 0 || (countryCompare == 0 && cityCompare <= 0);
                    if (pairIsSorted)
                    {
                        sortedPairs++;
                    }
                    else if (string.IsNullOrWhiteSpace(firstViolation))
                    {
                        firstViolation =
                            $"Hàng {row} ({P20GraderHelpers.NormalizeText(countryCurrent)} / {P20GraderHelpers.NormalizeText(cityCurrent)}) đứng trước hàng {row + 1} ({P20GraderHelpers.NormalizeText(countryNext)} / {P20GraderHelpers.NormalizeText(cityNext)}) không đúng thứ tự.";
                    }
                }

                if (totalPairs == 0)
                {
                    score += 8m;
                    result.Details.Add("Dữ liệu có 1 dòng nên mặc định đạt tiêu chí thứ tự sắp xếp.");
                }
                else
                {
                    score += Math.Round(8m * sortedPairs / totalPairs, 2, MidpointRounding.AwayFromZero);
                    if (sortedPairs == totalPairs)
                    {
                        result.Details.Add("Thứ tự dữ liệu thực tế đã đúng Country or region (A→Z), rồi City (A→Z).");
                    }
                    else
                    {
                        result.Errors.Add(
                            $"Thứ tự dữ liệu chưa được sắp xếp đúng hoàn toàn ({sortedPairs}/{totalPairs} cặp liên tiếp đúng). {firstViolation}");
                    }
                }

                var firstCountry = P20GraderHelpers.NormalizeText(worksheet.Cells[firstDataRow, countryCol].Text);
                var firstCity = P20GraderHelpers.NormalizeText(worksheet.Cells[firstDataRow, cityCol].Text);
                var lastCountry = P20GraderHelpers.NormalizeText(worksheet.Cells[lastDataRow, countryCol].Text);
                var lastCity = P20GraderHelpers.NormalizeText(worksheet.Cells[lastDataRow, cityCol].Text);

                var boundariesOk =
                    string.Equals(firstCountry, "Australia", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(firstCity, "Sydney", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(lastCountry, "United States", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(lastCity, "Seattle", StringComparison.OrdinalIgnoreCase)
                    && totalRows == 17;

                if (boundariesOk)
                {
                    score += 8m;
                    result.Details.Add("Biên dữ liệu sau sắp xếp là đúng (đầu bảng Australia/Sydney, cuối bảng United States/Seattle, đủ 17 dòng).");
                }
                else
                {
                    result.Errors.Add(
                        $"Biên hoặc số dòng dữ liệu sau sắp xếp chưa đúng. Đầu bảng='{firstCountry}/{firstCity}', cuối bảng='{lastCountry}/{lastCity}', số dòng={totalRows}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 3: {ex.Message}.");
            }

            return result;
        }
    }
}


