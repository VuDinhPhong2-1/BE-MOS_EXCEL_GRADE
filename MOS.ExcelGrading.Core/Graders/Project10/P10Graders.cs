using MOS.ExcelGrading.Core.Interfaces;
using MOS.ExcelGrading.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace MOS.ExcelGrading.Core.Graders.Project10
{
    public class P10T1Grader : ITaskGrader
    {
        public string TaskId => "P10-T1";
        public string TaskName => "Last semester: wrap text A3:F3";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Last semester");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Last semester'.");
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
                    result.Errors.Add($"Wrap Text chua dung tren A3:F3 ({wrappedCells}/6 o).");
                }

                var row3 = ws.Row(3);
                if (row3.Height > 15d)
                {
                    score += 1m;
                    result.Details.Add($"Chieu cao dong 3 hop le (Height={row3.Height:0.##}).");
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
                    result.Errors.Add("Tieu de tren dong 3 bi trong sau khi dinh dang.");
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

    public class P10T2Grader : ITaskGrader
    {
        public string TaskId => "P10-T2";
        public string TaskName => "Enrollment summary: named range Enrollment";
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
                decimal score = 0m;
                if (!P10GraderHelpers.TryGetDefinedName(studentSheet.Workbook, "Enrollment", out var definedValue))
                {
                    result.Errors.Add("Khong tim thay Named Range 'Enrollment'.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tim thay Named Range 'Enrollment'.");

                if (P10GraderHelpers.IsRangeMatch(definedValue, "A3:B7"))
                {
                    score += 2m;
                    result.Details.Add("Named Range 'Enrollment' dung vung A3:B7.");
                }
                else
                {
                    result.Errors.Add($"Named Range 'Enrollment' sai vung. Hien tai: '{definedValue}'.");
                }

                var normalizedDefinedValue = (definedValue ?? string.Empty)
                    .Replace(" ", string.Empty, StringComparison.Ordinal)
                    .Replace("'", string.Empty, StringComparison.Ordinal)
                    .ToUpperInvariant();
                if (normalizedDefinedValue.Contains("ENROLLMENTSUMMARY!", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Named Range dung sheet 'Enrollment summary'.");
                }
                else
                {
                    result.Errors.Add($"Named Range chua tro dung sheet 'Enrollment summary'. Hien tai: '{definedValue}'.");
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

    public class P10T3Grader : ITaskGrader
    {
        public string TaskId => "P10-T3";
        public string TaskName => "Income: table A3:B7, style Light14";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Income");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Income'.");
                    return result;
                }

                decimal score = 0m;
                var table = P10GraderHelpers.FindTableByAddress(ws, "A3:B7");
                if (table != null)
                {
                    score += 2m;
                    result.Details.Add("Table dung range A3:B7.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay table A3:B7. Hien tai: {P10GraderHelpers.JoinTableAddresses(ws)}.");
                    result.Score = score;
                    return result;
                }

                if (table.TableStyle == TableStyles.Light14)
                {
                    score += 1m;
                    result.Details.Add("Table style dung: TableStyleLight14.");
                }
                else
                {
                    result.Errors.Add($"Table style chua dung. Hien tai: {table.TableStyle}.");
                }

                var expectedHeaders = table.Columns.Count == 2
                                      && string.Equals(P10GraderHelpers.NormalizeText(table.Columns[0].Name), "Program", StringComparison.Ordinal)
                                      && string.Equals(P10GraderHelpers.NormalizeText(table.Columns[1].Name), "Total", StringComparison.Ordinal);
                if (expectedHeaders)
                {
                    score += 1m;
                    result.Details.Add("Ten cot table dung: Program, Total.");
                }
                else
                {
                    var columns = string.Join(", ", table.Columns.Select(c => c.Name));
                    result.Errors.Add($"Ten cot table chua dung. Hien tai: {columns}.");
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

    public class P10T4Grader : ITaskGrader
    {
        public string TaskId => "P10-T4";
        public string TaskName => "Last semester: remove Agriculture row from table";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Last semester");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Last semester'.");
                    return result;
                }

                decimal score = 0m;
                var table = P10GraderHelpers.FindTableByAddress(ws, "A3:F20");
                if (table != null)
                {
                    score += 2m;
                    result.Details.Add("Table 'Last semester' dung range A3:F20.");
                }
                else
                {
                    result.Errors.Add($"Khong tim thay table A3:F20. Hien tai: {P10GraderHelpers.JoinTableAddresses(ws)}.");
                    result.Score = score;
                    return result;
                }

                var agricultureRow = P10GraderHelpers.FindRowContainsText(
                    ws,
                    table.Address.Start.Column,
                    table.Address.Start.Row + 1,
                    table.Address.End.Row,
                    "Agriculture");
                if (agricultureRow < 0)
                {
                    score += 1m;
                    result.Details.Add("Khong con dong du lieu 'Agriculture' trong bang.");
                }
                else
                {
                    result.Errors.Add($"Van con dong 'Agriculture' tai hang {agricultureRow}.");
                }

                var dataRowCount = table.Address.Rows - 1;
                if (dataRowCount == 17 && ws.Tables.Count == 1)
                {
                    score += 1m;
                    result.Details.Add("So dong du lieu va so luong table hop le sau khi xoa.");
                }
                else
                {
                    result.Errors.Add($"So dong du lieu/table chua dung (rows={dataRowCount}, tables={ws.Tables.Count}).");
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

    public class P10T5Grader : ITaskGrader
    {
        public string TaskId => "P10-T5";
        public string TaskName => "Next semester: clustered column chart Program vs Average cost";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Next semester");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Next semester'.");
                    return result;
                }

                decimal score = 0m;
                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren sheet 'Next semester'.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tim thay chart tren sheet 'Next semester'.");

                if (chart.ChartType == eChartType.ColumnClustered)
                {
                    score += 1m;
                    result.Details.Add("Chart dung loai Clustered Column.");
                }
                else
                {
                    result.Errors.Add($"Loai chart chua dung. Hien tai: {chart.ChartType}.");
                }

                if (P10GraderHelpers.IsSeriesRangeMatch(chart, "A4:A21", "E4:E21"))
                {
                    score += 1m;
                    result.Details.Add("Series chart dung Program (A4:A21) va Average cost (E4:E21).");
                }
                else
                {
                    var series = chart.Series.FirstOrDefault();
                    result.Errors.Add($"Series chart chua dung. X='{series?.XSeries}', Y='{series?.Series}'.");
                }

                if (P10GraderHelpers.IsChartBoundsMatch(chart, "H3:O17"))
                {
                    score += 1m;
                    result.Details.Add("Vi tri chart hop le o ben phai bang (H3:O17).");
                }
                else
                {
                    result.Errors.Add($"Vi tri chart chua dung. Hien tai: {P10GraderHelpers.GetChartBounds(chart)}.");
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

    public class P10T6Grader : ITaskGrader
    {
        public string TaskId => "P10-T6";
        public string TaskName => "Enrollment summary: chart Style 7 + Monochromatic Palette 6";
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
                var ws = P10GraderHelpers.GetSheet(studentSheet.Workbook, "Enrollment summary");
                if (ws == null)
                {
                    result.Errors.Add("Khong tim thay sheet 'Enrollment summary'.");
                    return result;
                }

                decimal score = 0m;
                var chart = ws.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart == null)
                {
                    result.Errors.Add("Khong tim thay chart tren sheet 'Enrollment summary'.");
                    return result;
                }

                score += 1m;
                result.Details.Add("Tim thay chart tren sheet 'Enrollment summary'.");

                var (choiceStyle, fallbackStyle) = P10GraderHelpers.GetChartAlternateStyles(chart);
                var styleMatch = string.Equals(choiceStyle, "108", StringComparison.Ordinal)
                                 && string.Equals(fallbackStyle, "8", StringComparison.Ordinal);
                if (styleMatch)
                {
                    score += 1m;
                    result.Details.Add("Chart style XML hop le (Choice=108, Fallback=8), tuong ung Style 7.");
                }
                else
                {
                    result.Errors.Add($"Chart style XML chua dung. Choice='{choiceStyle}', Fallback='{fallbackStyle}'.");
                }

                var (colorStyleId, chartStyleId) = P10GraderHelpers.GetStyleManagerIds(chart);
                if (string.Equals(colorStyleId, "19", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Color style ID = 19 (Monochromatic Palette 6).");
                }
                else
                {
                    result.Errors.Add($"Color style ID chua dung. Hien tai: '{colorStyleId}' (mong doi: 19).");
                }

                if (string.Equals(chartStyleId, "268", StringComparison.Ordinal))
                {
                    score += 1m;
                    result.Details.Add("Chart style ID = 268 (Style 7 cho dang chart nay).");
                }
                else
                {
                    result.Errors.Add($"Chart style ID chua dung. Hien tai: '{chartStyleId}' (mong doi: 268).");
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
