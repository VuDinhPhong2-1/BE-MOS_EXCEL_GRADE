using System.Globalization;
using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Excel.Project13
{
    public class P13T1Grader : ITaskGrader
    {
        public string TaskId => "P13-T1";
        public string TaskName => "Shirt Orders: thay Amber bang Gold";
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
                var ws = P13GraderHelpers.GetSheet(studentSheet.Workbook, "Shirt Orders");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm th?y sheet 'Shirt Orders'.");
                    return result;
                }

                decimal score = 0m;
                var amberCount = 0;
                var goldCount = 0;
                for (var row = 6; row <= 199; row++)
                {
                    var color = (ws.Cells[row, 4].Text ?? string.Empty).Trim();
                    if (string.Equals(color, "Amber", StringComparison.OrdinalIgnoreCase))
                    {
                        amberCount++;
                    }

                    if (string.Equals(color, "Gold", StringComparison.Ordinal))
                    {
                        goldCount++;
                    }
                }

                if (amberCount == 0)
                {
                    score += 2m;
                    result.Details.Add("Không còn giá tr? 'Amber' trong c?t Shirt Color.");
                }
                else
                {
                    result.Errors.Add($"Van con {amberCount} giá tr? 'Amber' trong c?t Shirt Color.");
                }

                if (goldCount > 0)
                {
                    score += 2m;
                    result.Details.Add($"Da thay bang 'Gold' ({goldCount} o).");
                }
                else
                {
                    result.Errors.Add("Không tìm th?y giá tr? 'Gold' trong c?t Shirt Color.");
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




