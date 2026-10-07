using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class ExcelProject06ConditionsTests
    {
        private static object CreateOfficePackage(Dictionary<string, string> xmlParts)
        {
            var service = typeof(XmlGradingRuleService);
            var packageType = service.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var partsProp = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;
            foreach (var kvp in xmlParts)
            {
                partsProp[kvp.Key] = kvp.Value;
            }
            return package;
        }

        private static (bool IsPassed, string Message) InvokeEvaluator(string methodName, object? config, object package)
        {
            var service = typeof(XmlGradingRuleService);
            var method = service.GetMethod(methodName, BindingFlags.Static | BindingFlags.NonPublic)!;
            var outcome = method.Invoke(null, new object?[] { config, package })!;

            var isPassed = (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
            var message = (string?)outcome.GetType().GetProperty("Message")!.GetValue(outcome) ?? string.Empty;
            return (isPassed, message);
        }

        [Fact]
        public void EvaluateExcelTableTotalRow_WhenTotalsRowShown_Passes()
        {
            var tableXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<table xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" id=""1"" name=""Table1"" displayName=""Table1"" ref=""A1:E10"" totalsRowCount=""1"" totalsRowShown=""1"">
    <tableColumns count=""2"">
        <tableColumn id=""1"" name=""Item""/>
        <tableColumn id=""2"" name=""Entries"" totalsRowFunction=""count""/>
    </tableColumns>
</table>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/tables/table1.xml"] = tableXml
            });

            var config = new ExcelTableTotalRowConfig
            {
                RequireTotalRow = true,
                ColumnName = "Entries"
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelTableTotalRow", config, package);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void EvaluateExcelTableTotalRow_WhenNoTotalRow_Fails()
        {
            var tableXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<table xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" id=""1"" name=""Table1"" displayName=""Table1"" ref=""A1:E10"" totalsRowShown=""0"">
    <tableColumns count=""2"">
        <tableColumn id=""1"" name=""Item""/>
        <tableColumn id=""2"" name=""Entries""/>
    </tableColumns>
</table>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/tables/table1.xml"] = tableXml
            });

            var config = new ExcelTableTotalRowConfig
            {
                RequireTotalRow = true
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelTableTotalRow", config, package);
            Assert.False(isPassed);
        }

        [Fact]
        public void EvaluateExcelFormulaReferences_MultiCandidateCell_PassesWhenAnyCandidateMatches()
        {
            var worksheetXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <sheetData>
        <row r=""12"">
            <c r=""E12"">
                <f>MAX(E3:E10)</f>
                <v>500</v>
            </c>
        </row>
    </sheetData>
</worksheet>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/worksheets/sheet1.xml"] = worksheetXml
            });

            var config = new ExcelFormulaReferencesConfig
            {
                SourceFile = "xl/worksheets/sheet1.xml",
                Cell = "E10,E11,E12,E13,E14",
                RequiredFunctions = new List<string> { "MAX" }
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelFormulaReferences", config, package);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void EvaluateExcelChartType_2DPieWithFormulaFragments_Passes()
        {
            var chartXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
    <c:chart>
        <c:plotArea>
            <c:pieChart>
                <c:ser>
                    <c:cat>
                        <c:strRef>
                            <c:f>'Qtr 1'!$A$2:$A$10</c:f>
                        </c:strRef>
                    </c:cat>
                    <c:val>
                        <c:numRef>
                            <c:f>'Qtr 1'!Total</c:f>
                        </c:numRef>
                    </c:val>
                </c:ser>
            </c:pieChart>
        </c:plotArea>
    </c:chart>
</c:chartSpace>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/charts/chart1.xml"] = chartXml
            });

            var config = new ExcelChartTypeConfig
            {
                ExpectedChartType = "2-D Pie",
                ChartSourceFile = "xl/charts/chart1.xml",
                RequiredFormulaFragments = new List<string> { "Total" }
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelChartType", config, package);
            Assert.True(isPassed, message);
        }
        [Fact]
        public void EvaluateExcelChartType_PieChart_TableColumnMapping_Passes()
        {
            var chartXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
    <c:chart>
        <c:plotArea>
            <c:pieChart>
                <c:ser>
                    <c:tx>
                        <c:strRef>
                            <c:f>'Qtr 1'!$E$2</c:f>
                            <c:strCache>
                                <c:ptCount val=""1""/>
                                <c:pt idx=""0""><c:v>Total</c:v></c:pt>
                            </c:strCache>
                        </c:strRef>
                    </c:tx>
                    <c:cat>
                        <c:strRef>
                            <c:f>'Qtr 1'!$A$3:$A$10</c:f>
                        </c:strRef>
                    </c:cat>
                    <c:val>
                        <c:numRef>
                            <c:f>'Qtr 1'!$E$3:$E$10</c:f>
                        </c:numRef>
                    </c:val>
                </c:ser>
            </c:pieChart>
        </c:plotArea>
    </c:chart>
</c:chartSpace>";

            var tableXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<table xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" id=""1"" name=""Q1_Attendance"" displayName=""Q1_Attendance"" ref=""A2:E11"">
    <tableColumns count=""5"">
        <tableColumn id=""1"" name=""Entries""/>
        <tableColumn id=""2"" name=""Jan""/>
        <tableColumn id=""3"" name=""Feb""/>
        <tableColumn id=""4"" name=""Mar""/>
        <tableColumn id=""5"" name=""Total""/>
    </tableColumns>
</table>";

            var workbookXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <sheets>
        <sheet name=""Qtr 1"" sheetId=""1"" r:id=""rId1""/>
    </sheets>
</workbook>";

            var sheetRelsXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"" Target=""../tables/table1.xml""/>
    <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"" Target=""../drawings/drawing1.xml""/>
</Relationships>";

            var drawingXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <xdr:twoCellAnchor>
        <xdr:graphicFrame>
            <a:graphic>
                <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" r:id=""rId1""/>
                </a:graphicData>
            </a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>";

            var drawingRelsXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"" Target=""../charts/chart1.xml""/>
</Relationships>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/workbook.xml"] = workbookXml,
                ["xl/worksheets/sheet1.xml"] = @"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><tableParts count=""1""><tablePart xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" r:id=""rId1""/></tableParts></worksheet>",
                ["xl/worksheets/_rels/sheet1.xml.rels"] = sheetRelsXml,
                ["xl/tables/table1.xml"] = tableXml,
                ["xl/drawings/drawing1.xml"] = drawingXml,
                ["xl/drawings/_rels/drawing1.xml.rels"] = drawingRelsXml,
                ["xl/charts/chart1.xml"] = chartXml
            });

            var config = new ExcelChartTypeConfig
            {
                WorksheetName = "Qtr 1",
                ExpectedChartType = "2-D Pie",
                RequiredFormulaFragments = new List<string> { "Entries", "Total" }
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelChartType", config, package);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void EvaluateExcelChartType_WithPlacedBelowRow_PassesWhenBelowTable()
        {
            var workbookXml = @"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheets><sheet name=""Qtr 1"" sheetId=""1"" r:id=""rId1"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/></sheets></workbook>";
            var tableXml = @"<table xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" name=""Q1_Attendance"" displayName=""Q1_Attendance"" ref=""A2:E11""><tableColumns count=""5""><tableColumn id=""1"" name=""Entries""/><tableColumn id=""2"" name=""Col2""/><tableColumn id=""3"" name=""Col3""/><tableColumn id=""4"" name=""Col4""/><tableColumn id=""5"" name=""Total""/></tableColumns></table>";
            var chartXml = @"<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:chart><c:plotArea><c:pieChart><c:ser><c:tx><c:strRef><c:f>'Qtr 1'!$E$2</c:f><c:strCache><c:pt count=""1""><c:v>Total</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>'Qtr 1'!$A$3:$A$10</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'Qtr 1'!$E$3:$E$10</c:f></c:numRef></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>";
            var sheetRelsXml = @"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"" Target=""../tables/table1.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"" Target=""../drawings/drawing1.xml""/></Relationships>";
            // Chart starts at 0-indexed row 13 -> 1-based row 14
            var drawingXml = @"<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>13</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:graphicFrame><a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" r:id=""rId1""/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";
            var drawingRelsXml = @"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"" Target=""../charts/chart1.xml""/></Relationships>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/workbook.xml"] = workbookXml,
                ["xl/worksheets/sheet1.xml"] = @"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><tableParts count=""1""><tablePart xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" r:id=""rId1""/></tableParts></worksheet>",
                ["xl/worksheets/_rels/sheet1.xml.rels"] = sheetRelsXml,
                ["xl/tables/table1.xml"] = tableXml,
                ["xl/drawings/drawing1.xml"] = drawingXml,
                ["xl/drawings/_rels/drawing1.xml.rels"] = drawingRelsXml,
                ["xl/charts/chart1.xml"] = chartXml
            });

            var config = new ExcelChartTypeConfig
            {
                WorksheetName = "Qtr 1",
                ExpectedChartType = "2-D Pie",
                RequiredFormulaFragments = new List<string> { "Entries", "Total" },
                PlacedBelowRow = 11
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelChartType", config, package);
            Assert.True(isPassed, message);
            Assert.Contains("dòng 14", message);
        }

        [Fact]
        public void EvaluateExcelChartType_WithPlacedBelowRow_FailsWhenAboveOrInsideTable()
        {
            var workbookXml = @"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheets><sheet name=""Qtr 1"" sheetId=""1"" r:id=""rId1"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/></sheets></workbook>";
            var tableXml = @"<table xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" name=""Q1_Attendance"" displayName=""Q1_Attendance"" ref=""A2:E11""><tableColumns count=""5""><tableColumn id=""1"" name=""Entries""/><tableColumn id=""2"" name=""Col2""/><tableColumn id=""3"" name=""Col3""/><tableColumn id=""4"" name=""Col4""/><tableColumn id=""5"" name=""Total""/></tableColumns></table>";
            var chartXml = @"<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:chart><c:plotArea><c:pieChart><c:ser><c:tx><c:strRef><c:f>'Qtr 1'!$E$2</c:f><c:strCache><c:pt count=""1""><c:v>Total</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>'Qtr 1'!$A$3:$A$10</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'Qtr 1'!$E$3:$E$10</c:f></c:numRef></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>";
            var sheetRelsXml = @"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"" Target=""../tables/table1.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"" Target=""../drawings/drawing1.xml""/></Relationships>";
            // Chart starts at 0-indexed row 4 -> 1-based row 5 (inside table)
            var drawingXml = @"<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:graphicFrame><a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart""><c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" r:id=""rId1""/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";
            var drawingRelsXml = @"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"" Target=""../charts/chart1.xml""/></Relationships>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/workbook.xml"] = workbookXml,
                ["xl/worksheets/sheet1.xml"] = @"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><tableParts count=""1""><tablePart xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" r:id=""rId1""/></tableParts></worksheet>",
                ["xl/worksheets/_rels/sheet1.xml.rels"] = sheetRelsXml,
                ["xl/tables/table1.xml"] = tableXml,
                ["xl/drawings/drawing1.xml"] = drawingXml,
                ["xl/drawings/_rels/drawing1.xml.rels"] = drawingRelsXml,
                ["xl/charts/chart1.xml"] = chartXml
            });

            var config = new ExcelChartTypeConfig
            {
                WorksheetName = "Qtr 1",
                ExpectedChartType = "2-D Pie",
                RequiredFormulaFragments = new List<string> { "Entries", "Total" },
                PlacedBelowRow = 11
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelChartType", config, package);
            Assert.False(isPassed, message);
            Assert.Contains("chưa được đặt bên dưới dòng 11", message);
        }

        [Fact]
        public void EvaluateExcelChartType_UserDesktopFile_Passes()
        {
            var desktopPath = @"C:\Users\vup60\OneDrive\Máy tính\2019_Excel_106_BlogLog.xlsx";
            if (!System.IO.File.Exists(desktopPath))
            {
                return;
            }

            var xmlParts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var fileStream = new System.IO.FileStream(desktopPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
            using (var zip = new System.IO.Compression.ZipArchive(fileStream, System.IO.Compression.ZipArchiveMode.Read))
            {
                foreach (var entry in zip.Entries)
                {
                    if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                        entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    {
                        using var reader = new System.IO.StreamReader(entry.Open());
                        xmlParts[entry.FullName] = reader.ReadToEnd();
                    }
                }
            }

            var package = CreateOfficePackage(xmlParts);
            var config = new ExcelChartTypeConfig
            {
                WorksheetName = "Qtr 1",
                ExpectedChartType = "2-D Pie",
                RequiredFormulaFragments = new List<string> { "Entries", "Total" },
                PlacedBelowRow = 11
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelChartType", config, package);
            Assert.True(isPassed, message);
            Assert.Contains("dòng 16", message);

            var configTooFar = new ExcelChartTypeConfig
            {
                WorksheetName = "Qtr 1",
                ExpectedChartType = "2-D Pie",
                RequiredFormulaFragments = new List<string> { "Entries", "Total" },
                PlacedBelowRow = 20
            };

            var (isPassedTooFar, messageTooFar) = InvokeEvaluator("EvaluateExcelChartType", configTooFar, package);
            Assert.False(isPassedTooFar, messageTooFar);
            Assert.Contains("chưa được đặt bên dưới dòng 20", messageTooFar);
        }


        [Fact]
        public void EvaluateExcelTableCreate_ValidRangeAndStyle_Passes()
        {
            var tableXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<table xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" id=""1"" name=""Table1"" displayName=""Table1"" ref=""A2:E10"" headerRowCount=""1"">
    <tableColumns count=""5"">
        <tableColumn id=""1"" name=""Col1""/>
        <tableColumn id=""2"" name=""Col2""/>
        <tableColumn id=""3"" name=""Col3""/>
        <tableColumn id=""4"" name=""Col4""/>
        <tableColumn id=""5"" name=""Col5""/>
    </tableColumns>
    <tableStyleInfo name=""TableStyleLight14"" showRowStripes=""1""/>
</table>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/tables/table1.xml"] = tableXml
            });

            var config = new ExcelTableCreateConfig
            {
                ExpectedRange = "A2:E10",
                HasHeaderRow = true,
                ExpectedTableStyle = "Red, Table Style Light 14"
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelTableCreate", config, package);
            Assert.True(isPassed, message);
        }

        [Fact]
        public void EvaluateExcelChartQuickLayout_WithDataLabels_Passes()
        {
            var chartXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
    <c:chart>
        <c:plotArea>
            <c:barChart>
                <c:dLbls>
                    <c:showVal val=""1""/>
                    <c:dLblPos val=""outEnd""/>
                </c:dLbls>
            </c:barChart>
        </c:plotArea>
    </c:chart>
</c:chartSpace>";

            var package = CreateOfficePackage(new Dictionary<string, string>
            {
                ["xl/charts/chart1.xml"] = chartXml
            });

            var config = new ExcelChartQuickLayoutConfig
            {
                ChartSourceFile = "xl/charts/chart1.xml",
                LayoutNumber = 2,
                RequireDataLabels = true
            };

            var (isPassed, message) = InvokeEvaluator("EvaluateExcelChartQuickLayout", config, package);
            Assert.True(isPassed, message);
        }

    }
}

