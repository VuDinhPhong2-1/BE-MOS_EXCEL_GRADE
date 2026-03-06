using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T5Grader : ITaskGrader
    {
        public string TaskId => "P02-T5";
        public string TaskName => "Tao email tu First name voi @humongousinsurance.com";
        public decimal MaxScore => 4;

        public TaskResult Grade(ExcelWorksheet studentSheet, ExcelWorksheet answerSheet)
        {
            var result = new TaskResult
            {
                TaskId = TaskId,
                TaskName = TaskName,
                MaxScore = MaxScore
            };

            try
            {
                var contact = studentSheet.Workbook.Worksheets["Contact"];
                if (contact == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Contact'");
                    return result;
                }

                var table = contact.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay bang du lieu tren Contact");
                    return result;
                }

                var emailCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Email address", StringComparison.OrdinalIgnoreCase));
                var firstNameCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "First name", StringComparison.OrdinalIgnoreCase));

                if (emailCol == null || firstNameCol == null)
                {
                    result.Errors.Add("Thieu cot 'Email address' hoac 'First name' tren Contact");
                    return result;
                }

                var startRow = table.Address.Start.Row + 1;
                var endRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var emailColIndex = table.Address.Start.Column + emailCol.Position;
                var firstNameColIndex = table.Address.Start.Column + firstNameCol.Position;
                var totalRows = Math.Max(0, endRow - startRow + 1);
                if (totalRows == 0)
                {
                    result.Errors.Add("Khong co dong du lieu de cham");
                    return result;
                }

                var formulaRows = 0;
                var concatRows = 0;
                var strictDomainRows = 0;
                var firstNameRefRows = 0;

                for (var row = startRow; row <= endRow; row++)
                {
                    var rawFormula = contact.Cells[row, emailColIndex].Formula ?? string.Empty;
                    var formula = NormalizeFormula(rawFormula);
                    if (string.IsNullOrWhiteSpace(rawFormula))
                    {
                        continue;
                    }

                    formulaRows++;
                    if (formula.Contains("CONCAT(") || formula.Contains("CONCATENATE(") || formula.Contains("&"))
                    {
                        concatRows++;
                    }

                    var firstNameCellAddress = NormalizeFormula(contact.Cells[row, firstNameColIndex].Address);
                    if (formula.Contains("FIRSTNAME") ||
                        formula.Contains("FIRST NAME") ||
                        formula.Contains(firstNameCellAddress))
                    {
                        firstNameRefRows++;
                    }

                    // Task yeu cau chinh xac: "@humongousinsurance.com" (khong co dau cach truoc @).
                    var hasExactDomainLiteral =
                        rawFormula.Contains("\"@humongousinsurance.com\"", StringComparison.OrdinalIgnoreCase);

                    // Fallback theo ket qua hien thi: email phai dung dang FirstName@domain.
                    var firstName = contact.Cells[row, firstNameColIndex].Text ?? string.Empty;
                    var expectedEmail = $"{firstName}@humongousinsurance.com";
                    var actualEmail = contact.Cells[row, emailColIndex].Text ?? string.Empty;
                    var hasExpectedOutput = string.Equals(
                        actualEmail,
                        expectedEmail,
                        StringComparison.OrdinalIgnoreCase);

                    if (hasExactDomainLiteral || hasExpectedOutput)
                    {
                        strictDomainRows++;
                    }
                }

                decimal score = 0;
                score += ComputeComponentScore(formulaRows, totalRows, 1m, result,
                    "Cot Email da co cong thuc cho tat ca dong",
                    $"Thieu cong thuc cot Email ({formulaRows}/{totalRows})");
                score += ComputeComponentScore(concatRows, totalRows, 1m, result,
                    "Cong thuc co phep ghep chuoi",
                    $"Cong thuc ghep chuoi chua day du ({concatRows}/{totalRows})");
                score += ComputeComponentScore(strictDomainRows, totalRows, 1m, result,
                    "Cong thuc da chen dung domain @humongousinsurance.com",
                    $"Domain email chua dung chuan (có thể đã bị sai chính tả) ({strictDomainRows}/{totalRows})");
                score += ComputeComponentScore(firstNameRefRows, totalRows, 1m, result,
                    "Cong thuc tham chieu First name",
                    $"Tham chieu First name chua dung/du ({firstNameRefRows}/{totalRows})");

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static decimal ComputeComponentScore(
            int passedRows,
            int totalRows,
            decimal maxComponent,
            TaskResult result,
            string successMessage,
            string errorMessage)
        {
            if (totalRows <= 0) return 0;

            if (passedRows == totalRows)
            {
                result.Details.Add(successMessage);
                return maxComponent;
            }

            result.Errors.Add(errorMessage);
            return Math.Round((decimal)passedRows / totalRows * maxComponent, 2);
        }

        private static string NormalizeFormula(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return string.Empty;
            }

            return formula
                .Replace("=", string.Empty)
                .Replace("$", string.Empty)
                .Replace(" ", string.Empty)
                .ToUpperInvariant();
        }
    }
}
