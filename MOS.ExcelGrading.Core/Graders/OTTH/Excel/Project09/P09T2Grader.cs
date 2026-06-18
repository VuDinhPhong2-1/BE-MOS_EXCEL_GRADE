using OfficeOpenXml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project09
{
    public class P09T2Grader : ITaskGrader
    {
        public string TaskId => "P09-T2";
        public string TaskName => "Huy tron o A1, ap dung kieu Title, co chu 24, Bold";
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
                var cell = studentSheet.Cells["A1"];
                decimal score = 0;

                // Rule 1: Không merge (1 di?m)
                if (!cell.Merge)
                {
                    score += 1;
                    result.Details.Add("? Ðã h?y merge cell A1");
                }
                else
                {
                    result.Errors.Add("? Cell A1 v?n còn merge");
                }

                // Rule 2: Style = Title (1 di?m)
                if (cell.StyleName.Contains("Title", StringComparison.OrdinalIgnoreCase))
                {
                    score += 1;
                    result.Details.Add("? Áp d?ng Title style");
                }
                else
                {
                    result.Errors.Add($"? Style không dúng (hi?n t?i: {cell.StyleName})");
                }

                // Rule 3: Font size = 24 (1 di?m)
                if (cell.Style.Font.Size == 24)
                {
                    score += 1;
                    result.Details.Add("? Font size 24pt");
                }
                else
                {
                    result.Errors.Add($"? Font size sai (hi?n t?i: {cell.Style.Font.Size}pt)");
                }

                // Rule 4: Bold (1 di?m)
                if (cell.Style.Font.Bold)
                {
                    score += 1;
                    result.Details.Add("? In d?m");
                }
                else
                {
                    result.Errors.Add("? Chua in d?m");
                }

                result.Score = score;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"? L?i: {ex.Message}");
            }

            return result;
        }
    }
}

// minor-sync: non-functional graders update




