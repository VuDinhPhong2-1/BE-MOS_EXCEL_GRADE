using OfficeOpenXml;
using OfficeOpenXml.Style;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Project02
{
    public class P02T1Grader : ITaskGrader
    {
        public string TaskId => "P02-T1";
        public string TaskName => "Canh trai + thut le cot Agent trong New Policy";
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
                    result.Errors.Add("Khong tim thay bang du lieu tren sheet New Policy");
                    return result;
                }

                var agentColumn = table.Columns.FirstOrDefault(c =>
                    string.Equals(c.Name?.Trim(), "Agent", StringComparison.OrdinalIgnoreCase));
                if (agentColumn == null)
                {
                    result.Errors.Add("Khong tim thay cot 'Agent'");
                    return result;
                }

                var dataStartRow = table.Address.Start.Row + 1;
                var dataEndRow = table.Address.End.Row - (table.ShowTotal ? 1 : 0);
                var colIndex = table.Address.Start.Column + agentColumn.Position;

                var totalRows = Math.Max(0, dataEndRow - dataStartRow + 1);
                if (totalRows == 0)
                {
                    result.Errors.Add("Khong co dong du lieu de cham cot Agent");
                    return result;
                }

                var leftAlignedRows = 0;
                var indentOneRows = 0;

                for (var row = dataStartRow; row <= dataEndRow; row++)
                {
                    var cell = ws.Cells[row, colIndex];
                    if (cell.Style.HorizontalAlignment == ExcelHorizontalAlignment.Left)
                    {
                        leftAlignedRows++;
                    }

                    if (cell.Style.Indent == 1)
                    {
                        indentOneRows++;
                    }
                }

                decimal score = 0;
                score += 1m; // Tim dung cot Agent

                if (leftAlignedRows == totalRows)
                {
                    score += 1.5m;
                    result.Details.Add($"Tat ca o Agent da canh trai ({leftAlignedRows}/{totalRows})");
                }
                else
                {
                    result.Errors.Add($"Canh trai chua day du ({leftAlignedRows}/{totalRows})");
                }

                if (indentOneRows == totalRows)
                {
                    score += 1.5m;
                    result.Details.Add($"Tat ca o Agent co indent = 1 ({indentOneRows}/{totalRows})");
                }
                else
                {
                    result.Errors.Add($"Indent level = 1 chua day du ({indentOneRows}/{totalRows})");
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

