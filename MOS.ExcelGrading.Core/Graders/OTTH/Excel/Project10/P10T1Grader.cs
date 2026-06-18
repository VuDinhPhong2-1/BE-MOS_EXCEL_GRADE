using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project10
{
    public class P10T1Grader : ITaskGrader
    {
        public string TaskId => "P10-T1";
        public string TaskName => "Last semester: bat Wrap Text cho A3:F3";
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
                var wrappedCells = 0;
                for (var col = 1; col <= 6; col++)
                {
                    if (ws.Cells[3, col].Style.WrapText)
                    {
                        wrappedCells++;
                    }
                }

                if (wrappedCells == 6)
                {
                    score += 2m;
                    result.Details.Add("Da bat Wrap Text cho day du A3:F3.");
                }
                else
                {
                    result.Errors.Add($"Wrap Text chua dúng tren A3:F3 ({wrappedCells}/6 o).");
                }

                var row3 = ws.Row(3);
                if (row3.Height > 15d)
                {
                    score += 1m;
                    result.Details.Add($"Chieu cao dòng 3 hop le (Height={row3.Height:0.##}).");
                }
                else
                {
                    result.Errors.Add($"Dong 3 chua duoc dan chieu cao de hien thi wrap text (Height={row3.Height:0.##}).");
                }

                var hasHeaderText = !string.IsNullOrWhiteSpace(ws.Cells["A3"].Text)
                                    && !string.IsNullOrWhiteSpace(ws.Cells["F3"].Text);
                if (hasHeaderText)
                {
                    score += 1m;
                    result.Details.Add("Noi dung tieu de A3/F3 van duoc giu.");
                }
                else
                {
                    result.Errors.Add("Tieu de tren dòng 3 bi trong sau khi dinh dang.");
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




