using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using System.Globalization;

namespace MOS.ExcelGrading.Core.Graders.Project04
{
    public class P04T7Grader : ITaskGrader
    {
        public string TaskId => "P04-T7";
        public string TaskName => "Sort Classes theo Instructor A-Z, sau đó Section giảm dần";
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
                    result.Errors.Add("Không tìm thấy sheet Classes");
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
                    result.Details.Add($"Đọc đủ dữ liệu để sort ({rows.Count} dòng)");
                }
                else
                {
                    result.Errors.Add($"Không đủ dòng để kiểm tra sort ({rows.Count}/21)");
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
                    result.Details.Add("Đã sort Instructor tăng dần (A-Z)");
                }
                else
                {
                    result.Errors.Add("Thứ tự Instructor chưa tăng dần A-Z");
                }

                if (sectionSortedInGroup)
                {
                    result.Score += 1.5m;
                    result.Details.Add("Trong từng Instructor, Section được sort giảm dần");
                }
                else
                {
                    result.Errors.Add("Section chưa giảm dần trong nhóm Instructor");
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


// minor-sync: non-functional graders update
