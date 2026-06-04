using OfficeOpenXml;
using OfficeOpenXml.Table;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T3Grader : ITaskGrader
    {
        public string TaskId => "P02-T3";
        public string TaskName => "Thêm Total Row và tổng theo tháng trong bảng New Policy";
        public decimal MaxScore => 4;

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
                var ws = studentSheet.Workbook.Worksheets["New Policy"];
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'New Policy'");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Không tìm thấy bảng dữ liệu trên sheet 'New Policy'");
                    return result;
                }

                decimal score = 0;
                if (table.ShowTotal)
                {
                    score += 1m;
                    result.Details.Add("Đã bật Total Row");
                }
                else
                {
                    result.Errors.Add("Chưa bật Total Row cho bảng");
                }

                var monthCols = new[] { "January", "February", "March", "April", "May", "June" };
                var monthSumCols = 0;
                foreach (var colName in monthCols)
                {
                    var col = table.Columns.FirstOrDefault(c =>
                        string.Equals(c.Name?.Trim(), colName, StringComparison.OrdinalIgnoreCase));
                    if (col == null)
                    {
                        continue;
                    }

                    if (col.TotalsRowFunction == RowFunctions.Sum)
                    {
                        monthSumCols++;
                    }
                }

                if (monthSumCols == monthCols.Length)
                {
                    score += 2m;
                    result.Details.Add("Tổng theo 6 tháng đã cấu hình bằng hàm SUM");
                }
                else
                {
                    result.Errors.Add($"Tổng theo tháng chưa đầy đủ ({monthSumCols}/6 cột tháng)");
                }

                var totalCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Total", StringComparison.OrdinalIgnoreCase));
                if (totalCol != null && totalCol.TotalsRowFunction == RowFunctions.Sum)
                {
                    score += 1m;
                    result.Details.Add("Cột Total đã tổng cộng bằng SUM");
                }
                else
                {
                    result.Errors.Add("Cột Total chưa cấu hình tổng cộng đúng");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}


// minor-sync: non-functional graders update
