using OfficeOpenXml;
using OfficeOpenXml.Table;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T3Grader : ITaskGrader
    {
        public string TaskId => "P02-T3";
        public string TaskName => "Them Total Row va tong theo thang trong bang New Policy";
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

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay bang du lieu tren New Policy");
                    return result;
                }

                decimal score = 0;
                if (table.ShowTotal)
                {
                    score += 1m;
                    result.Details.Add("Da bat Total Row");
                }
                else
                {
                    result.Errors.Add("Chua bat Total Row cho bang");
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
                    result.Details.Add("Tong theo 6 thang da cau hinh bang ham SUM");
                }
                else
                {
                    result.Errors.Add($"Tong theo thang chua day du ({monthSumCols}/6 cot thang)");
                }

                var totalCol = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Total", StringComparison.OrdinalIgnoreCase));
                if (totalCol != null && totalCol.TotalsRowFunction == RowFunctions.Sum)
                {
                    score += 1m;
                    result.Details.Add("Cot Total da tong cong bang SUM");
                }
                else
                {
                    result.Errors.Add("Cot Total chua cau hinh tong cong dung");
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

