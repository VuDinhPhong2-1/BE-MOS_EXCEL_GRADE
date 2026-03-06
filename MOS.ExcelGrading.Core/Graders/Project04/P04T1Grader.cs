using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T1Grader : ITaskGrader
    {
        public string TaskId => "P04-T1";
        public string TaskName => "Import Substitutes data + table style Medium 1";
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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Substitutes");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet Substitutes");
                    return result;
                }

                var table = ws.Tables.FirstOrDefault();
                if (table == null)
                {
                    result.Errors.Add("Khong tim thay table tren sheet Substitutes");
                    return result;
                }

                result.Score += 1m;
                result.Details.Add($"Tim thay table '{table.Name}'");

                var addrOk = P04GraderHelpers.NormalizeAddress(table.Address.Address) == "A4:D10";
                if (addrOk && table.Address.Rows == 7)
                {
                    result.Score += 1m;
                    result.Details.Add("Table dung vung A4:D10");
                }
                else
                {
                    result.Errors.Add($"Table sai vung. Hien tai: {table.Address.Address}");
                }

                if (table.TableStyle == OfficeOpenXml.Table.TableStyles.Medium1)
                {
                    result.Score += 1m;
                    result.Details.Add("Da ap dung table style Medium 1");
                }
                else
                {
                    result.Errors.Add($"Table style chua dung Medium 1 (hien tai: {table.TableStyle})");
                }

                var expectedHeaders = new[] { "Name", "First name", "Object", "Preferred contact" };
                var headerOk = expectedHeaders
                    .Select((h, idx) => string.Equals(
                        (table.Columns.ElementAtOrDefault(idx)?.Name ?? string.Empty).Trim(),
                        h,
                        StringComparison.OrdinalIgnoreCase))
                    .All(x => x);

                var sampleOk = string.Equals(ws.Cells["A5"].Text.Trim(), "Syed", StringComparison.OrdinalIgnoreCase)
                               && ws.Cells["D10"].Text.Contains("@", StringComparison.Ordinal);

                if (headerOk && sampleOk)
                {
                    result.Score += 1m;
                    result.Details.Add("Header va du lieu mau dung voi file Substitutes");
                }
                else
                {
                    result.Errors.Add("Header hoac du lieu mau tren Substitutes chua dung");
                }

                result.Score = Math.Min(MaxScore, result.Score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Loi: {ex.Message}");
            }

            return result;
        }
    }
}

