using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project08
{
    public class P08T3Grader : ITaskGrader
    {
        public string TaskId => "P08-T3";
        public string TaskName => "Xoa thong tin ca nhan trong Document Properties";
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
                var props = studentSheet.Workbook.Properties;
                var checks = new (string Label, string Value)[]
                {
                    ("Author", props.Author ?? string.Empty),
                    ("LastModifiedBy", props.LastModifiedBy ?? string.Empty),
                    ("Manager", props.Manager ?? string.Empty),
                    ("Company", props.Company ?? string.Empty),
                    ("Title", props.Title ?? string.Empty),
                    ("Subject", props.Subject ?? string.Empty),
                    ("Keywords", props.Keywords ?? string.Empty),
                    ("Category", props.Category ?? string.Empty),
                    ("Comments", props.Comments ?? string.Empty),
                };

                var emptyCount = 0;
                foreach (var check in checks)
                {
                    if (string.IsNullOrWhiteSpace(check.Value))
                    {
                        emptyCount++;
                    }
                    else
                    {
                        result.Errors.Add($"{check.Label} van con du lieu: '{check.Value}'.");
                    }
                }

                if (emptyCount == checks.Length)
                {
                    result.Details.Add("Da xoa cac thuoc tinh thong tin ca nhan chinh.");
                }

                result.Score = Math.Round(MaxScore * emptyCount / checks.Length, 2);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}
