using System.Xml;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;

namespace MOS.ExcelGrading.Core.Graders.Project20
{
    public class P20T2Grader : ITaskGrader
    {
        public string TaskId => "P20-T2";
        public string TaskName => "Trong trang tính “London”, xóa tất cả các quy tắc định dạng có điều kiện.";
        public decimal MaxScore => 17m;

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
                var worksheet = P20GraderHelpers.GetSheet(studentSheet.Workbook, "London");
                if (worksheet == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'London'.");
                    return result;
                }

                decimal score = 0m;

                var ruleCount = worksheet.ConditionalFormatting.Count;
                if (ruleCount == 0)
                {
                    score += 10m;
                    result.Details.Add("Sheet London không còn quy tắc Conditional Formatting nào trong bộ sưu tập quy tắc.");
                }
                else
                {
                    var summary = string.Join(
                        "; ",
                        worksheet.ConditionalFormatting.Select(rule => $"{rule.Type} @ {rule.Address?.Address}"));
                    result.Errors.Add($"Sheet London vẫn còn {ruleCount} quy tắc Conditional Formatting: {summary}.");
                }

                var xml = worksheet.WorksheetXml;
                var ns = new XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var cfNodes = xml.SelectNodes("/x:worksheet/x:conditionalFormatting", ns);
                var xmlRuleGroups = cfNodes?.Count ?? 0;
                if (xmlRuleGroups == 0)
                {
                    score += 7m;
                    result.Details.Add("Worksheet XML của London không còn node conditionalFormatting.");
                }
                else
                {
                    var sqrefs = new List<string>();
                    if (cfNodes != null)
                    {
                        foreach (XmlNode node in cfNodes)
                        {
                            var sqref = node.Attributes?["sqref"]?.Value ?? string.Empty;
                            sqrefs.Add(string.IsNullOrWhiteSpace(sqref) ? "(không có sqref)" : sqref);
                        }
                    }

                    result.Errors.Add(
                        $"Worksheet XML của London vẫn còn {xmlRuleGroups} nhóm conditionalFormatting, sqref: {string.Join(", ", sqrefs)}.");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi khi chấm Task 2: {ex.Message}.");
            }

            return result;
        }
    }
}

