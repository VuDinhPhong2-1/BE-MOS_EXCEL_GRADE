using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project18
{
    public class P18T5Grader : ITaskGrader
    {
        public string TaskId => "P18-T5";
        public string TaskName => "Trong trang tính \"Contact\", tại cột Email Address, sử dụng hàm để tạo địa chỉ email cho mỗi người bằng cách kết hợp First Name với “@woodgrovebank.com”.";
        public decimal MaxScore => 20m;

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
                var worksheet = P18GraderHelpers.GetSheet(studentSheet.Workbook, "Contact");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Contact'.");
                    return result;
                }

                decimal score = 0m;

                var table = P18GraderHelpers.FindTableByHeaders(worksheet, "First name", "Email Address");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu có cột First name và Email Address.");
                    return result;
                }

                score += 4m;
                result.Details.Add("Đã tìm thấy bảng dữ liệu Contact với đúng cột First name và Email Address.");

                var columnsByNormalizedName = table.Columns
                    .ToDictionary(
                        column => P18GraderHelpers.NormalizeIdentifier(column.Name),
                        column => table.Address.Start.Column + column.Position,
                        StringComparer.OrdinalIgnoreCase);

                if (!columnsByNormalizedName.TryGetValue("FIRSTNAME", out var firstNameColumn)
                    || !columnsByNormalizedName.TryGetValue("EMAILADDRESS", out var emailColumn))
                {
                    result.Errors.Add("Không xác định được đúng vị trí cột First name hoặc Email Address.");
                    result.Score = score;
                    return result;
                }

                var firstDataRow = table.Address.Start.Row + 1;
                var lastDataRow = table.Address.End.Row;
                var totalRows = Math.Max(0, lastDataRow - firstDataRow + 1);
                if (totalRows == 0)
                {
                    result.Errors.Add("Bảng Contact không có dòng dữ liệu để chấm.");
                    result.Score = score;
                    return result;
                }

                var formulaPresentCount = 0;
                var formulaLogicCount = 0;
                var outputCorrectCount = 0;
                const string expectedDomain = "@woodgrovebank.com";

                for (var row = firstDataRow; row <= lastDataRow; row++)
                {
                    var emailCell = worksheet.Cells[row, emailColumn];
                    var firstNameCell = worksheet.Cells[row, firstNameColumn];
                    var rawFormula = emailCell.Formula ?? string.Empty;
                    var normalizedFormula = P18GraderHelpers.NormalizeFormula(rawFormula);

                    if (!string.IsNullOrWhiteSpace(rawFormula))
                    {
                        formulaPresentCount++;

                        var hasConcatOperator = normalizedFormula.Contains("CONCAT(", StringComparison.Ordinal)
                                                || normalizedFormula.Contains("CONCATENATE(", StringComparison.Ordinal)
                                                || normalizedFormula.Contains("&", StringComparison.Ordinal);
                        var firstNameAddress = ExcelCellBase.GetAddress(row, firstNameColumn).ToUpperInvariant();
                        var hasFirstNameReference = normalizedFormula.Contains("FIRSTNAME", StringComparison.Ordinal)
                                                    || normalizedFormula.Contains(firstNameAddress, StringComparison.Ordinal);
                        var hasExactDomainLiteral = P18GraderHelpers.HasExactQuotedLiteral(rawFormula, expectedDomain);
                        var hasSpaceBeforeDomain = P18GraderHelpers.HasSpaceBeforeDomainLiteral(rawFormula, expectedDomain);

                        if (hasConcatOperator && hasFirstNameReference && hasExactDomainLiteral && !hasSpaceBeforeDomain)
                        {
                            formulaLogicCount++;
                        }
                    }

                    var firstNameText = (firstNameCell.Text ?? string.Empty).Trim();
                    var expectedEmail = $"{firstNameText}{expectedDomain}";
                    var actualEmail = (emailCell.Text ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(firstNameText)
                        && string.Equals(actualEmail, expectedEmail, StringComparison.OrdinalIgnoreCase))
                    {
                        outputCorrectCount++;
                    }
                }

                score += Math.Round(4m * formulaPresentCount / totalRows, 2, MidpointRounding.AwayFromZero);
                if (formulaPresentCount == totalRows)
                {
                    result.Details.Add("Cột Email Address đã có công thức cho toàn bộ dòng dữ liệu.");
                }
                else
                {
                    result.Errors.Add(
                        $"Cột Email Address chưa có công thức đầy đủ ({formulaPresentCount}/{totalRows} dòng).");
                }

                score += Math.Round(6m * formulaLogicCount / totalRows, 2, MidpointRounding.AwayFromZero);
                if (formulaLogicCount == totalRows)
                {
                    result.Details.Add("Công thức Email Address ghép đúng First name với chuỗi \"@woodgrovebank.com\" và không có khoảng trắng sai.");
                }
                else
                {
                    result.Errors.Add(
                        $"Công thức Email Address chưa đúng logic ghép chuỗi hoặc sai domain ở một số dòng ({formulaLogicCount}/{totalRows} dòng đúng).");
                }

                score += Math.Round(6m * outputCorrectCount / totalRows, 2, MidpointRounding.AwayFromZero);
                if (outputCorrectCount == totalRows)
                {
                    result.Details.Add("Kết quả email hiển thị đúng chuẩn FirstName@woodgrovebank.com cho toàn bộ dòng.");
                }
                else
                {
                    result.Errors.Add(
                        $"Kết quả email hiển thị chưa đúng ở một số dòng ({outputCorrectCount}/{totalRows} dòng đúng).");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 5: {ex.Message}.");
            }

            return result;
        }
    }
}

