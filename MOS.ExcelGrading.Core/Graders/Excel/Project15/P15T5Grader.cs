using System.Xml;
using System.Text.RegularExpressions;
using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project15
{
    public class P15T5Grader : ITaskGrader
    {
        public string TaskId => "P15-T5";
        public string TaskName => "Customers: hoan thien cot CurrentAge khong lam doi dinh dang";
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
                var ws = P15GraderHelpers.GetSheet(studentSheet.Workbook, "Customers");
                if (ws == null)
                {
                    result.Errors.Add("Không tìm thấy sheet 'Customers'.");
                    return result;
                }

                if (!P15GraderHelpers.TryFindColumnByHeader(
                    ws,
                    P15GraderHelpers.IsCurrentAgeColumnName,
                    out var headerRow,
                    out var currentAgeCol))
                {
                    result.Errors.Add("Không tìm thấy cột 'CurrentAge'.");
                    return result;
                }

                var dataStart = headerRow + 1;
                var dataEnd = ws.Dimension.End.Row;
                if (P15GraderHelpers.TryFindColumnByHeader(ws, P15GraderHelpers.IsIdColumnName, out _, out var idCol))
                {
                    dataEnd = P15GraderHelpers.GetLastDataRowInColumn(ws, idCol, dataStart);
                }

                if (dataStart <= 0 || dataEnd < dataStart)
                {
                    result.Errors.Add("Không xác định được vung dữ liệu cột CurrentAge.");
                    return result;
                }

                var totalRows = dataEnd - dataStart + 1;
                var hasBirthDateCol = P15GraderHelpers.TryFindColumnByHeader(
                    ws,
                    P15GraderHelpers.IsBirthDateColumnName,
                    out _,
                    out var birthDateCol);

                decimal score = 0m;
                var filledCount = 0;
                var seriesCorrectCount = 0;
                var styleIds = new List<int>(capacity: totalRows);

                for (var row = dataStart; row <= dataEnd; row++)
                {
                    var ageCell = ws.Cells[row, currentAgeCol];
                    styleIds.Add(ageCell.StyleID);

                    var hasFormulaProp = !string.IsNullOrWhiteSpace(ageCell.Formula);
                    var hasFormulaNode = P15GraderHelpers.TryGetCellFormulaNode(ws, row, currentAgeCol, out var formulaNode);
                    var hasValue = !string.IsNullOrWhiteSpace(ageCell.Text);
                    if (hasValue || hasFormulaProp || hasFormulaNode)
                    {
                        filledCount++;
                    }

                    var formulaText = hasFormulaProp
                        ? ageCell.Formula
                        : (hasFormulaNode ? (formulaNode?.InnerText ?? string.Empty) : string.Empty);
                    var normalizedFormula = P15GraderHelpers.NormalizeFormula(formulaText);
                    var isFormulaRow = hasFormulaProp || hasFormulaNode;

                    var isSeriesCorrect = false;
                    if (isFormulaRow)
                    {
                        if (string.IsNullOrWhiteSpace(normalizedFormula))
                        {
                            // Shared-formula follower cells can have empty inner text in XML.
                            isSeriesCorrect = true;
                        }
                        else if (hasBirthDateCol)
                        {
                            var birthCellAddress = ExcelCellBase.GetAddress(row, birthDateCol);
                            var birthRef = P15GraderHelpers.NormalizeFormula(birthCellAddress);
                            isSeriesCorrect =
                                normalizedFormula.Contains("YEAR(TODAY())", StringComparison.Ordinal)
                                && normalizedFormula.Contains($"YEAR({birthRef})", StringComparison.Ordinal);
                        }
                    }
                    else if (hasBirthDateCol && hasValue)
                    {
                        // Accept pre-calculated values if they match YEAR(TODAY())-YEAR(BirthDate).
                        var birthValue = ws.Cells[row, birthDateCol].Value;
                        DateTime birthDate;
                        if (birthValue is DateTime dt)
                        {
                            birthDate = dt;
                        }
                        else if (birthValue is double oa)
                        {
                            birthDate = DateTime.FromOADate(oa);
                        }
                        else if (!DateTime.TryParse(ws.Cells[row, birthDateCol].Text, out birthDate))
                        {
                            birthDate = DateTime.MinValue;
                        }

                        if (birthDate != DateTime.MinValue && int.TryParse(ageCell.Text, out var ageValue))
                        {
                            var expectedAge = DateTime.Today.Year - birthDate.Year;
                            isSeriesCorrect = ageValue == expectedAge;
                        }
                    }

                    if (isSeriesCorrect)
                    {
                        seriesCorrectCount++;
                    }
                }

                if (filledCount == totalRows)
                {
                    score += 2m;
                    result.Details.Add($"Cot CurrentAge da duoc dien day du ({filledCount}/{totalRows} dòng).");
                }
                else
                {
                    result.Errors.Add($"Cot CurrentAge chua dien du. Hiện tại: {filledCount}/{totalRows} dòng.");
                }

                if (seriesCorrectCount == totalRows)
                {
                    score += 1m;
                    result.Details.Add("Chuoi dữ liệu CurrentAge dung (cong thuc/ket qua hop le tren tat ca dòng).");
                }
                else
                {
                    result.Errors.Add($"Chuoi dữ liệu CurrentAge chưa đúng ({seriesCorrectCount}/{totalRows} dòng hop le).");
                }

                var innerRows = Enumerable.Range(dataStart + 1, Math.Max(0, totalRows - 2)).ToList();
                var oddStyles = innerRows
                    .Where(r => (r - dataStart) % 2 == 1)
                    .Select(r => ws.Cells[r, currentAgeCol].StyleID)
                    .Distinct()
                    .ToList();
                var evenStyles = innerRows
                    .Where(r => (r - dataStart) % 2 == 0)
                    .Select(r => ws.Cells[r, currentAgeCol].StyleID)
                    .Distinct()
                    .ToList();
                var styleOk = styleIds.All(id => id > 0)
                              && oddStyles.Count <= 1
                              && evenStyles.Count <= 1
                              && (oddStyles.Count == 0 || evenStyles.Count == 0 || oddStyles[0] != evenStyles[0]);
                if (styleOk)
                {
                    score += 1m;
                    result.Details.Add("Định dạng cột CurrentAge duoc giu on dinh theo mau dòng.");
                }
                else
                {
                    result.Errors.Add("Định dạng cột CurrentAge co dau hieu bi thay doi (khong con on dinh theo mau dòng).");
                }

                result.Score = Math.Min(MaxScore, score);
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Lỗi: {ex.Message}");
            }

            return result;
        }
    }
}




