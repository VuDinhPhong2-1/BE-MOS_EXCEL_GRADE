using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using System.Globalization;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T7Grader : ITaskGrader
    {
        public string TaskId => "P04-T7";
        public string TaskName => "Sort Classes theo Instructor A-Z, sau do Section giam dan";
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
                var ws = P04GraderHelpers.GetSheet(studentSheet, "Classes");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet Classes");
                    return result;
                }

                var rows = new List<(string Instructor, int Section, int Row)>();
                for (var r = 5; r <= 25; r++)
                {
                    var instructor = ws.Cells[r, 6].Text.Trim(); // col F
                    var sectionText = ws.Cells[r, 2].Text.Trim(); // col B
                    if (string.IsNullOrWhiteSpace(instructor) || !P04GraderHelpers.TryParseSection(sectionText, out var section))
                    {
                        continue;
                    }

                    rows.Add((instructor, section, r));
                }

                if (rows.Count >= 20)
                {
                    result.Score += 1m;
                    result.Details.Add($"Doc du du lieu de sort ({rows.Count} dong)");
                }
                else
                {
                    result.Errors.Add($"Khong du dong de kiem tra sort ({rows.Count}/21)");
                    return result;
                }

                var instructorSorted = true;
                var sectionSortedInGroup = true;
                for (var i = 1; i < rows.Count; i++)
                {
                    var prev = rows[i - 1];
                    var curr = rows[i];
                    var cmp = string.Compare(prev.Instructor, curr.Instructor, CultureInfo.InvariantCulture, CompareOptions.IgnoreCase);
                    if (cmp > 0)
                    {
                        instructorSorted = false;
                        break;
                    }

                    if (cmp == 0 && prev.Section < curr.Section)
                    {
                        sectionSortedInGroup = false;
                        break;
                    }
                }

                if (instructorSorted)
                {
                    result.Score += 1.5m;
                    result.Details.Add("Da sort Instructor tang dan (A-Z)");
                }
                else
                {
                    result.Errors.Add("Thu tu Instructor chua tang dan A-Z");
                }

                if (sectionSortedInGroup)
                {
                    result.Score += 1.5m;
                    result.Details.Add("Trong tung Instructor, Section duoc sort giam dan");
                }
                else
                {
                    result.Errors.Add("Section chua giam dan trong nhom Instructor");
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

