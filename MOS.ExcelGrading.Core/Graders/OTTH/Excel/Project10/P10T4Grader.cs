using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project10
{
    public class P10T4Grader : ITaskGrader
    {
        public string TaskId => "P10-T4";
        public string TaskName => "Last semester: xoa dong Agriculture khoi Table";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Last semester");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Last semester'.");
                    return result;
                }

                decimal score = 0m;
                var table = P10GraderHelpers.FindTableByAddress(ws, "A3:F20");
                if (table != null)
                {
                    score += 2m;
                    result.Details.Add("Table 'Last semester' dung range A3:F20.");
                }
                else
                {
                    result.Errors.Add($"Không tìm th?y table A3:F20. Hi?n t?i: {P10GraderHelpers.JoinTableAddresses(ws)}.");
                    result.Score = score;
                    return result;
                }

                var agricultureRow = P10GraderHelpers.FindRowContainsText(
                    ws,
                    table.Address.Start.Column,
                    table.Address.Start.Row + 1,
                    table.Address.End.Row,
                    "Agriculture");
                if (agricultureRow < 0)
                {
                    score += 1m;
                    result.Details.Add("Không còn dòng d? li?u 'Agriculture' trong bang.");
                }
                else
                {
                    result.Errors.Add($"Van con dòng 'Agriculture' tai hàng {agricultureRow}.");
                }

                var dataRowCount = table.Address.Rows - 1;
                if (dataRowCount == 17 && ws.Tables.Count == 1)
                {
                    score += 1m;
                    result.Details.Add("So dòng d? li?u va so luong table hop le sau khi xoa.");
                }
                else
                {
                    result.Errors.Add($"So dòng d? li?u/table chua dúng (rows={dataRowCount}, tables={ws.Tables.Count}).");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"L?i: {ex.Message}");
            }

            return result;
        }
    }
}




