using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T2Grader : ITaskGrader
    {
        public string TaskId => "P02-T2";
        public string TaskName => "Chen sparkline Win/Loss tai J5:J13";
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
                var ws = studentSheet.Workbook.Worksheets["New Policy"];
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'New Policy'");
                    return result;
                }

                if (ws.SparklineGroups.Count == 0)
                {
                    result.Errors.Add("Chua co sparkline group nao tren sheet New Policy");
                    return result;
                }

                decimal score = 0;
                score += 1m;

                var targetGroup = ws.SparklineGroups.First();

                // EPPlus enum in-memory cho Win/Loss thuong la Stacked.
                var sparklineTypeText = targetGroup.Type.ToString();
                if (sparklineTypeText.Contains("Stacked", StringComparison.OrdinalIgnoreCase) ||
                    sparklineTypeText.Contains("WinLoss", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1m;
                    result.Details.Add($"Loai sparkline hop le: {sparklineTypeText}");
                }
                else
                {
                    result.Errors.Add($"Loai sparkline chua dung Win/Loss (hien tai: {sparklineTypeText})");
                }

                var locationRange = GetAddressLikeString(targetGroup, "LocationAddress", "LocationRange");
                if (NormalizeAddress(locationRange) == "J5:J13")
                {
                    score += 1m;
                    result.Details.Add("Location range dung J5:J13");
                }
                else
                {
                    result.Errors.Add($"Location range chua dung (hien tai: {locationRange})");
                }

                var expectedRows = Enumerable.Range(5, 9).ToList();
                var validSparklineRows = 0;
                var correctDataRangeRows = 0;

                foreach (var sp in targetGroup.Sparklines)
                {
                    var cellAddress = GetAddressLikeString(sp, "Cell");
                    var dataAddress = GetAddressLikeString(sp, "RangeAddress", "Range");
                    if (!TryGetRowFromAddress(cellAddress, out var row))
                    {
                        continue;
                    }

                    if (expectedRows.Contains(row))
                    {
                        validSparklineRows++;
                        var normalizedData = NormalizeAddress(dataAddress);
                        var expectedData = $"B{row}:G{row}";
                        if (normalizedData == expectedData)
                        {
                            correctDataRangeRows++;
                        }
                    }
                }

                if (validSparklineRows == expectedRows.Count && correctDataRangeRows == expectedRows.Count)
                {
                    score += 1m;
                    result.Details.Add("Tat ca 9 sparkline dung data range tu thang 1 den thang 6");
                }
                else
                {
                    result.Errors.Add(
                        $"Sparkline row/data range chua day du (rows={validSparklineRows}/9, data={correctDataRangeRows}/9)");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }

        private static string GetAddressLikeString(object? obj, params string[] propertyNames)
        {
            if (obj == null) return string.Empty;

            var type = obj.GetType();
            foreach (var name in propertyNames)
            {
                var prop = type.GetProperty(name);
                if (prop == null) continue;

                var value = prop.GetValue(obj);
                if (value == null) continue;

                var addrProp = value.GetType().GetProperty("Address");
                if (addrProp != null)
                {
                    var addrValue = addrProp.GetValue(value)?.ToString();
                    if (!string.IsNullOrWhiteSpace(addrValue))
                    {
                        return addrValue;
                    }
                }

                var asText = value.ToString();
                if (!string.IsNullOrWhiteSpace(asText))
                {
                    return asText;
                }
            }

            return string.Empty;
        }

        private static string NormalizeAddress(string? address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return string.Empty;
            }

            var text = address.Replace("$", string.Empty).Trim();
            var excl = text.LastIndexOf('!');
            if (excl >= 0 && excl + 1 < text.Length)
            {
                text = text[(excl + 1)..];
            }

            return text.Trim('\'');
        }

        private static bool TryGetRowFromAddress(string address, out int row)
        {
            row = 0;
            var normalized = NormalizeAddress(address);
            var digits = new string(normalized.Where(char.IsDigit).ToArray());
            return int.TryParse(digits, out row);
        }
    }
}

