using System;
using System.Collections.Generic;
using System.Reflection;
using MOS.ExcelGrading.Core.Models;
using MOS.ExcelGrading.Core.Services;
using Xunit;

namespace MOS.ExcelGrading.Word.UnitTests
{
    public class PptConditionTests
    {
        private static object CreateOfficePackage(Dictionary<string, string> xmlParts, List<string>? attachedFileNames = null)
        {
            var serviceType = typeof(XmlGradingRuleService);
            var packageType = serviceType.GetNestedType("OfficePackage", BindingFlags.NonPublic)!;
            var package = Activator.CreateInstance(packageType, nonPublic: true)!;
            var parts = (Dictionary<string, string>)packageType.GetProperty("XmlParts")!.GetValue(package)!;

            foreach (var kvp in xmlParts)
            {
                parts[kvp.Key] = kvp.Value;
            }

            if (attachedFileNames != null)
            {
                packageType.GetProperty("AttachedFileNames")!.SetValue(package, attachedFileNames.ToArray());
            }

            return package;
        }

        private static (bool IsPassed, string Message) InvokeEvaluator(string methodName, object config, object package)
        {
            var serviceType = typeof(XmlGradingRuleService);
            var method = serviceType.GetMethod(methodName, BindingFlags.Static | BindingFlags.NonPublic)!;
            var outcome = method.Invoke(null, new object?[] { config, package })!;

            var isPassed = (bool)outcome.GetType().GetProperty("IsPassed")!.GetValue(outcome)!;
            var message = (string?)outcome.GetType().GetProperty("Message")!.GetValue(outcome) ?? string.Empty;
            return (isPassed, message);
        }

        [Fact]
        public void PptExportedFile_StrictCaseSensitive_MatchesExact_Passes()
        {
            var config = new PptExportedFileConfig
            {
                ExpectedFileName = "Backcountry.pdf",
                CaseSensitive = true
            };

            var package = CreateOfficePackage(new Dictionary<string, string>(), new List<string> { "Backcountry.pdf" });
            var (isPassed, message) = InvokeEvaluator("EvaluatePptExportedFile", config, package);

            Assert.True(isPassed, message);
        }

        [Fact]
        public void PptExportedFile_StrictCaseSensitive_WrongCasing_Fails()
        {
            var config = new PptExportedFileConfig
            {
                ExpectedFileName = "Backcountry.pdf",
                CaseSensitive = true
            };

            var package = CreateOfficePackage(new Dictionary<string, string>(), new List<string> { "backcountry.pdf" });
            var (isPassed, message) = InvokeEvaluator("EvaluatePptExportedFile", config, package);

            Assert.False(isPassed, message);
            Assert.Contains("Backcountry.pdf", message);
        }

        [Fact]
        public void PptVideo_SummarySlideContainsVideo_Passes()
        {
            var config = new PptVideoConfig
            {
                Slide = new PptSlideRef { SlideIndex = 2 },
                RequireVideoOnly = true
            };

            const string presXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<p:presentation xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""
                xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <p:sldIdLst>
        <p:sldId id=""256"" r:id=""rId1""/>
        <p:sldId id=""257"" r:id=""rId2""/>
    </p:sldIdLst>
</p:presentation>";

            const string presRelsXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide1.xml""/>
    <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide2.xml""/>
</Relationships>";

            const string slide2Xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<p:sld xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""
       xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <p:cSld>
        <p:spTree>
            <p:pic>
                <p:nvPicPr>
                    <p:cNvPr id=""5"" name=""Screen Recording 1""/>
                    <p:nvPr>
                        <a:videoFile r:link=""rIdMedia1""/>
                    </p:nvPr>
                </p:nvPicPr>
            </p:pic>
        </p:spTree>
    </p:cSld>
</p:sld>";

            var parts = new Dictionary<string, string>
            {
                ["ppt/presentation.xml"] = presXml,
                ["ppt/_rels/presentation.xml.rels"] = presRelsXml,
                ["ppt/slides/slide1.xml"] = @"<p:sld xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""><p:cSld><p:spTree/></p:cSld></p:sld>",
                ["ppt/slides/slide2.xml"] = slide2Xml
            };

            var package = CreateOfficePackage(parts);
            var (isPassed, message) = InvokeEvaluator("EvaluatePptVideo", config, package);

            Assert.True(isPassed, message);
        }

        [Fact]
        public void PptPictureCropShape_EllipseGeometry_Passes()
        {
            var config = new PptPictureCropShapeConfig
            {
                Slide = new PptSlideRef { SlideIndex = 1 },
                ExpectedShapePreset = "ellipse"
            };

            const string presXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<p:presentation xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""
                xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <p:sldIdLst>
        <p:sldId id=""256"" r:id=""rId1""/>
    </p:sldIdLst>
</p:presentation>";

            const string presRelsXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide1.xml""/>
</Relationships>";

            const string slide1Xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<p:sld xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""
       xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
    <p:cSld>
        <p:spTree>
            <p:pic>
                <p:spPr>
                    <a:prstGeom prst=""ellipse""/>
                </p:spPr>
            </p:pic>
        </p:spTree>
    </p:cSld>
</p:sld>";

            var parts = new Dictionary<string, string>
            {
                ["ppt/presentation.xml"] = presXml,
                ["ppt/_rels/presentation.xml.rels"] = presRelsXml,
                ["ppt/slides/slide1.xml"] = slide1Xml
            };

            var package = CreateOfficePackage(parts);
            var (isPassed, message) = InvokeEvaluator("EvaluatePptPictureCropShape", config, package);

            Assert.True(isPassed, message);
        }

        [Fact]
        public void PptChartLegend_TopPosition_Passes()
        {
            var config = new PptChartLegendConfig
            {
                ExpectedPosition = "t"
            };

            const string chartXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
    <c:chart>
        <c:legend>
            <c:legendPos val=""t""/>
        </c:legend>
    </c:chart>
</c:chartSpace>";

            var parts = new Dictionary<string, string>
            {
                ["ppt/charts/chart1.xml"] = chartXml
            };

            var package = CreateOfficePackage(parts);
            var (isPassed, message) = InvokeEvaluator("EvaluatePptChartLegend", config, package);

            Assert.True(isPassed, message);
        }

        [Fact]
        public void PptMarkAsFinal_CompletedDocument_Passes()
        {
            var config = new PptMarkAsFinalConfig
            {
                ExpectedMarkAsFinal = true
            };

            const string customPropsXml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/custom-properties""
            xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
    <property fmtid=""{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"" pid=""2"" name=""_MarkAsFinal"">
        <vt:bool>true</vt:bool>
    </property>
</Properties>";

            var parts = new Dictionary<string, string>
            {
                ["docProps/custom.xml"] = customPropsXml
            };

            var package = CreateOfficePackage(parts);
            var (isPassed, message) = InvokeEvaluator("EvaluatePptMarkAsFinal", config, package);

            Assert.True(isPassed, message);
        }

        [Fact]
        public void PptRuleSet_LoadPptGm2RuleSet_Loads7ProjectsAnd35Tasks()
        {
            var serviceType = typeof(XmlGradingRuleService);
            var loadMethod = serviceType.GetMethod("LoadPptGm2RuleSet", BindingFlags.Static | BindingFlags.NonPublic)!;
            var ruleSet = (GradingRuleSet)loadMethod.Invoke(null, null)!;

            Assert.NotNull(ruleSet);
            Assert.Equal("ppt", ruleSet.Subject);
            Assert.Equal("GM2-2019", ruleSet.Version);
            Assert.Equal(7, ruleSet.Projects.Count);

            int totalTasks = 0;
            decimal totalScore = 0m;
            foreach (var project in ruleSet.Projects)
            {
                totalTasks += project.Tasks.Count;
                totalScore += project.MaxScore;
            }

            Assert.Equal(35, totalTasks);
            Assert.True(totalScore >= 999m && totalScore <= 1001m, $"Total score was {totalScore}");
        }

        [Fact]
        public void PptRuleSet_ValidateRuleSet_HasZeroErrors()
        {
            var serviceType = typeof(XmlGradingRuleService);
            var loadMethod = serviceType.GetMethod("LoadPptGm2RuleSet", BindingFlags.Static | BindingFlags.NonPublic)!;
            var ruleSet = (GradingRuleSet)loadMethod.Invoke(null, null)!;

            var validateMethod = serviceType.GetMethod("ValidateRuleSet", BindingFlags.Static | BindingFlags.NonPublic)!;
            var validation = (XmlRuleValidationResult)validateMethod.Invoke(null, new object[] { ruleSet })!;

            Assert.True(validation.IsValid, $"Validation failed with errors: {string.Join("; ", validation.Errors)}");
            Assert.Empty(validation.Errors);
        }

        private static string? FindSampleFile(string fileName)
        {
            var dir = new System.IO.DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null)
            {
                var target = System.IO.Path.Combine(dir.FullName, "sample_pptx", fileName);
                if (System.IO.File.Exists(target)) return target;
                dir = dir.Parent;
            }
            return null;
        }

        private static object CreateOfficePackageFromPptx(string pptxPath)
        {
            var xmlParts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var archive = System.IO.Compression.ZipFile.OpenRead(pptxPath))
            {
                foreach (var entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                        entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    {
                        using var stream = entry.Open();
                        using var reader = new System.IO.StreamReader(stream);
                        xmlParts[entry.FullName] = reader.ReadToEnd();
                    }
                }
            }
            return CreateOfficePackage(xmlParts);
        }

        [Fact]
        public void PptShapeArrange_BackcountryToursSample_SlideTitleWithEllipsis_Passes()
        {
            var samplePath = FindSampleFile("Project 10 - BackcountryTours.pptx");
            Assert.NotNull(samplePath);

            var package = CreateOfficePackageFromPptx(samplePath!);
            var config = new PptShapeArrangeConfig
            {
                Slide = new PptSlideRef { SlideTitle = "Our Rock Crawling adventures..." },
                CheckAlignMiddle = true,
                CheckBringToFront = true
            };

            var (isPassed, message) = InvokeEvaluator("EvaluatePptShapeArrange", config, package);
            Assert.True(isPassed, message);
        }
    }
}
