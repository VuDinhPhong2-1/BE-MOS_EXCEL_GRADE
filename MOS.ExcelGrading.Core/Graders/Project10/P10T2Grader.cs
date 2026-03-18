using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T2Grader : ITaskGrader
    {
        public string TaskId => "P10-T2";
        public string TaskName => "Last semester: table range and style";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Last semester");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Last semester'.");
                    return result;
                }

                decimal score = 0m;
                var table = P10GraderHelpers.FindTableByAddress(ws, "A3:F20");
                if (table != null)
                {
                    score += 2m;
                    result.Details.Add("Table dung range A3:F20.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay table A3:F20. Hien tai: {P10GraderHelpers.JoinTableAddresses(ws)}");
                }

                if (table != null
                    && table.TableStyle == TableStyles.Light14)
                {
                    score += 1m;
                    result.Details.Add("Table style dung: TableStyleLight14.");
                }
                else if (table != null)
                {
                    result.Errors.Add($"Table style chua dung. Hien tai: {table.TableStyle}.");
                }

                if (ws.Tables.Count == 1)
                {
                    score += 1m;
                    result.Details.Add("So luong table tren sheet dung (1 table).");
                }
                else
                {
                    result.Errors.Add($"So luong table tren sheet chua dung. Hien tai: {ws.Tables.Count}.");
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
