using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project18
{
    public class P18T3Grader : ITaskGrader
    {
        public string TaskId => "P18-T3";
        public string TaskName => "Trong trang tính \"New Accounts\", xóa hàng khỏi bảng có chứa dữ liệu Tailspin Toys. Không thay đổi bất kỳ nội dung nào bên ngoài bảng.";
        public decimal MaxScore => 25m;

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
                var worksheet = P18GraderHelpers.GetSheet(studentSheet.Workbook, "New Accounts");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'New Accounts'.");
                    return result;
                }

                decimal score = 0m;

                var table = P18GraderHelpers.FindTableByHeaders(
                    worksheet,
                    "Account",
                    "Opening Balance",
                    "Current Balance");
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu chứa các cột Account, Opening Balance và Current Balance.");
                    return result;
                }

                score += 5m;
                result.Details.Add("Đã tìm thấy đúng bảng dữ liệu cần thao tác trên sheet 'New Accounts'.");

                var expectedTableAddress = "A3:C10";
                var actualTableAddress = P18GraderHelpers.NormalizeRange(table.Address.Address);
                if (string.Equals(actualTableAddress, P18GraderHelpers.NormalizeRange(expectedTableAddress), StringComparison.OrdinalIgnoreCase))
                {
                    score += 6m;
                    result.Details.Add("Địa chỉ bảng sau khi xóa dòng là A3:C10.");
                }
                else
                {
                    result.Errors.Add(
                        $"Địa chỉ bảng sau khi xóa dòng chưa đúng. Hiện tại: {table.Address.Address}, mong đợi: {expectedTableAddress}.");
                }

                var accounts = new List<string>();
                for (var row = table.Address.Start.Row + 1; row <= table.Address.End.Row; row++)
                {
                    var accountText = (worksheet.Cells[row, table.Address.Start.Column].Text ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(accountText))
                    {
                        accounts.Add(accountText);
                    }
                }

                var normalizedAccounts = accounts
                    .Select(P18GraderHelpers.NormalizeIdentifier)
                    .ToList();

                var hasTailspin = normalizedAccounts.Any(account =>
                    account.Contains("TAILSPIN", StringComparison.OrdinalIgnoreCase));

                var expectedOrder = new[]
                {
                    "ADVENTUREWORKS",
                    "CAFEFOURTH",
                    "FABRIKAMINC",
                    "LAMNAHEALTHCARE",
                    "LUCERNEEDITIONS",
                    "RELECLOUD",
                    "PROSEWAREINC"
                };

                var orderMatched = normalizedAccounts.Count == expectedOrder.Length
                                   && normalizedAccounts.SequenceEqual(expectedOrder, StringComparer.OrdinalIgnoreCase);

                if (!hasTailspin && orderMatched)
                {
                    score += 8m;
                    result.Details.Add("Dòng chứa 'Tailspin Toys' đã được xóa đúng và thứ tự dữ liệu trong bảng vẫn chính xác.");
                }
                else
                {
                    if (hasTailspin)
                    {
                        result.Errors.Add("Vẫn còn dữ liệu 'Tailspin Toys' trong bảng New Accounts.");
                    }

                    if (!orderMatched)
                    {
                        result.Errors.Add(
                            $"Dữ liệu tài khoản trong bảng chưa đúng sau thao tác xóa. Hiện tại: {string.Join(" | ", accounts)}.");
                    }
                }

                var outsideMarkers = new List<(string Address, string Text)>();
                for (var row = 2; row <= 16; row++)
                {
                    for (var col = 26; col <= 31; col++) // Z:AE
                    {
                        var text = (worksheet.Cells[row, col].Text ?? string.Empty).Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            outsideMarkers.Add((ExcelCellBase.GetAddress(row, col), text));
                        }
                    }
                }

                var toysMarkerAddresses = outsideMarkers
                    .Where(marker => string.Equals(marker.Text, "Toys", StringComparison.OrdinalIgnoreCase))
                    .Select(marker => marker.Address)
                    .ToList();

                if (toysMarkerAddresses.Count == 0)
                {
                    score += 6m;
                    result.Details.Add("Không có marker ngoài bảng dạng 'Toys' trong vùng Z2:AE16, nên bỏ qua kiểm tra vị trí marker.");
                }
                else if (toysMarkerAddresses.Any(address => string.Equals(address, "AA6", StringComparison.OrdinalIgnoreCase)))
                {
                    score += 6m;
                    result.Details.Add("Nội dung ngoài bảng vẫn được giữ nguyên (marker 'Toys' vẫn ở ô AA6).");
                }
                else
                {
                    result.Errors.Add(
                        $"Marker 'Toys' ngoài bảng bị lệch vị trí. Hiện tại ở: {string.Join(", ", toysMarkerAddresses)}; mong đợi: AA6.");
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
