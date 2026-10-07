using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Services
{
    public partial class XmlGradingRuleService
    {
        private static readonly XNamespace PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace RelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace OfficeRelsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace P14Ns = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private static readonly XNamespace P15Ns = "http://schemas.microsoft.com/office/powerpoint/2012/main";
        private static readonly XNamespace P166Ns = "http://schemas.microsoft.com/office/powerpoint/2018/4/main";
        private static readonly XNamespace CNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private static readonly XNamespace CxNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        private static readonly XNamespace DgmNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

        public const long EmuPerInch = 914400;
        public const long EmuPerCm = 360000;

        private static bool IsPptSpecialConditionSupported(string specialConditionType)
        {
            return string.Equals(specialConditionType, SpecialConditionTypes.PptPictureCropShape, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptShapeSize, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptShapeGroup, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptChartLegend, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSmartArt, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptComment, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSlideTitles, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptVideo, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptTable, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSection, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptPictureStyle, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptShapeArrange, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSummaryZoom, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptExportedFile, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptMasterPicture, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSlideTransition, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptAnimation, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptMarkAsFinal, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptPrintSettings, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptTextColumns, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptNotesMasterPlaceholders, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSlideSize, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptChartType, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptAltText, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptHyperlink, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptTextBox, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSlideLayout, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptShapeStyle, StringComparison.OrdinalIgnoreCase)
                || string.Equals(specialConditionType, SpecialConditionTypes.PptSlideBackground, StringComparison.OrdinalIgnoreCase);
        }

        private static void CollectPptRequiredOfficeParts(RequiredOfficeParts requiredParts)
        {
            requiredParts.XmlParts.Add("ppt/presentation.xml");
            requiredParts.XmlParts.Add("ppt/_rels/presentation.xml.rels");
            requiredParts.XmlParts.Add("ppt/presProps.xml");
            requiredParts.XmlParts.Add("docProps/core.xml");
            requiredParts.XmlParts.Add("docProps/app.xml");
            requiredParts.XmlParts.Add("docProps/custom.xml");

            requiredParts.XmlPartPrefixes.Add("ppt/slides");
            requiredParts.XmlPartPrefixes.Add("ppt/slideLayouts");
            requiredParts.XmlPartPrefixes.Add("ppt/slideMasters");
            requiredParts.XmlPartPrefixes.Add("ppt/notesMasters");
            requiredParts.XmlPartPrefixes.Add("ppt/notesSlides");
            requiredParts.XmlPartPrefixes.Add("ppt/charts");
            requiredParts.XmlPartPrefixes.Add("ppt/diagrams");
            requiredParts.XmlPartPrefixes.Add("ppt/comments");
        }

        private sealed record PptSlideItem(int Index, string Path, XDocument Doc);

        private static SpecialConditionEvalOutcome FailSpecialCondition(string message, string? fixAction)
        {
            var fullMessage = string.IsNullOrWhiteSpace(fixAction)
                ? message
                : $"{message} (Gợi ý: {fixAction})";
            return new SpecialConditionEvalOutcome { IsPassed = false, Message = fullMessage };
        }

        private static List<PptSlideItem> ResolveOrderedSlides(OfficePackage package)
        {
            var result = new List<PptSlideItem>();
            if (!package.XmlParts.TryGetValue("ppt/presentation.xml", out var presXml))
            {
                return result;
            }

            XDocument presDoc;
            try
            {
                presDoc = XDocument.Parse(presXml);
            }
            catch
            {
                return result;
            }

            var relsMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (package.XmlParts.TryGetValue("ppt/_rels/presentation.xml.rels", out var presRelsXml))
            {
                try
                {
                    var relsDoc = XDocument.Parse(presRelsXml);
                    foreach (var rel in relsDoc.Descendants(RelsNs + "Relationship"))
                    {
                        var id = rel.Attribute("Id")?.Value;
                        var target = rel.Attribute("Target")?.Value;
                        if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                        {
                            var normalizedTarget = target.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase)
                                ? target
                                : "ppt/" + target.TrimStart('/');
                            relsMap[id] = NormalizeSourceFile(normalizedTarget);
                        }
                    }
                }
                catch
                {
                }
            }

            var sldIds = presDoc.Descendants(PNs + "sldId").ToList();
            int slideCounter = 1;

            if (sldIds.Count > 0 && relsMap.Count > 0)
            {
                foreach (var sldId in sldIds)
                {
                    var rId = sldId.Attribute(OfficeRelsNs + "id")?.Value
                        ?? sldId.Attribute(XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships") + "id")?.Value;

                    if (!string.IsNullOrWhiteSpace(rId) && relsMap.TryGetValue(rId, out var targetPath))
                    {
                        if (package.XmlParts.TryGetValue(targetPath, out var slideXml))
                        {
                            try
                            {
                                var slideDoc = XDocument.Parse(slideXml);
                                result.Add(new PptSlideItem(slideCounter++, targetPath, slideDoc));
                            }
                            catch
                            {
                            }
                        }
                    }
                }
            }

            // Fallback nếu không parse được qua sldIdLst
            if (result.Count == 0)
            {
                var slidePaths = package.XmlParts.Keys
                    .Where(k => k.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && k.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(k => k, StringComparer.OrdinalIgnoreCase);

                foreach (var path in slidePaths)
                {
                    try
                    {
                        var doc = XDocument.Parse(package.XmlParts[path]);
                        result.Add(new PptSlideItem(slideCounter++, path, doc));
                    }
                    catch
                    {
                    }
                }
            }

            return result;
        }

        private static bool TryResolveTargetSlide(
            OfficePackage package,
            PptSlideRef? slideRef,
            out PptSlideItem slideInfo,
            out string? errorMessage)
        {
            slideInfo = null!;
            errorMessage = null;

            var slides = ResolveOrderedSlides(package);
            if (slides.Count == 0)
            {
                errorMessage = "Không tìm thấy trang chiếu (slide) nào trong tệp PowerPoint.";
                return false;
            }

            // 1. Chỉ định slideIndex
            if (slideRef?.SlideIndex.HasValue == true)
            {
                int index = slideRef.SlideIndex.Value;
                if (index == -1) // Slide cuối
                {
                    slideInfo = slides.Last();
                    return true;
                }

                if (index < 1 || index > slides.Count)
                {
                    errorMessage = $"Không tìm thấy Slide {index} (Bài thuyết trình có {slides.Count} slide).";
                    return false;
                }

                slideInfo = slides[index - 1];
                return true;
            }

            // 2. Chỉ định slideTitle
            if (!string.IsNullOrWhiteSpace(slideRef?.SlideTitle))
            {
                var titleTarget = NormalizeSlideTitle(slideRef.SlideTitle);
                foreach (var slide in slides)
                {
                    var title = NormalizeSlideTitle(GetSlideTitle(slide.Doc));
                    if (!string.IsNullOrWhiteSpace(title) && (title.Contains(titleTarget, StringComparison.OrdinalIgnoreCase) || titleTarget.Contains(title, StringComparison.OrdinalIgnoreCase)))
                    {
                        slideInfo = slide;
                        return true;
                    }

                    // Fallback: Tìm trong toàn bộ text trên slide
                    var allText = NormalizePlainText(GetAllSlideText(slide.Doc)).Replace("…", "...");
                    if (!string.IsNullOrWhiteSpace(allText) && allText.Contains(titleTarget, StringComparison.OrdinalIgnoreCase))
                    {
                        slideInfo = slide;
                        return true;
                    }
                }

                errorMessage = $"Không tìm thấy slide có tiêu đề chứa '{slideRef.SlideTitle}'.";
                return false;
            }

            // 3. Chỉ định anchorText
            if (!string.IsNullOrWhiteSpace(slideRef?.AnchorText))
            {
                var anchor = NormalizePlainText(slideRef.AnchorText);
                foreach (var slide in slides)
                {
                    var text = GetAllSlideText(slide.Doc);
                    if (NormalizePlainText(text).Contains(anchor, StringComparison.OrdinalIgnoreCase))
                    {
                        slideInfo = slide;
                        return true;
                    }
                }

                errorMessage = $"Không tìm thấy slide chứa nội dung '{slideRef.AnchorText}'.";
                return false;
            }

            // Mặc định lấy slide 1 nếu không chỉ định
            slideInfo = slides.First();
            return true;
        }

        private static string NormalizeSlideTitle(string? title)
        {
            if (string.IsNullOrWhiteSpace(title)) return string.Empty;
            return NormalizePlainText(title)
                .Replace("…", "...")
                .TrimEnd('.', ' ', '…')
                .Trim();
        }

        private static string GetSlideTitle(XDocument slideDoc)
        {
            var titleSp = slideDoc.Descendants(PNs + "sp")
                .FirstOrDefault(sp =>
                {
                    var ph = sp.Descendants(PNs + "ph").FirstOrDefault();
                    var type = ph?.Attribute("type")?.Value;
                    return type == "title" || type == "ctrTitle";
                });

            if (titleSp != null)
            {
                return GetElementText(titleSp);
            }

            // Hoặc lấy văn bản của shape đầu tiên có văn bản
            var textSp = slideDoc.Descendants(PNs + "sp")
                .FirstOrDefault(sp => sp.Descendants(ANs + "t").Any(t => !string.IsNullOrWhiteSpace(t.Value)));
            return textSp != null ? GetElementText(textSp) : string.Empty;
        }

        private static string GetElementText(XElement element)
        {
            var texts = element.Descendants(ANs + "t").Select(t => t.Value);
            return string.Join(" ", texts);
        }

        private static string GetAllSlideText(XDocument slideDoc)
        {
            var texts = slideDoc.Descendants(ANs + "t").Select(t => t.Value);
            return string.Join(" ", texts);
        }

        private static List<XElement> FindSlideShapes(XDocument slideDoc, PptShapeSelector? selector)
        {
            var spTree = slideDoc.Descendants(PNs + "spTree").FirstOrDefault();
            if (spTree == null) return new List<XElement>();

            var allShapes = spTree.Elements()
                .Where(e => e.Name == PNs + "sp" || e.Name == PNs + "pic" || e.Name == PNs + "grpSp" || e.Name == PNs + "graphicFrame")
                .ToList();

            if (selector == null) return allShapes;

            var filtered = allShapes.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(selector.ShapeType))
            {
                if (selector.ShapeType.Equals("picture", StringComparison.OrdinalIgnoreCase) || selector.ShapeType.Equals("pic", StringComparison.OrdinalIgnoreCase))
                    filtered = filtered.Where(e => e.Name == PNs + "pic");
                else if (selector.ShapeType.Equals("group", StringComparison.OrdinalIgnoreCase))
                    filtered = filtered.Where(e => e.Name == PNs + "grpSp");
                else if (selector.ShapeType.Equals("shape", StringComparison.OrdinalIgnoreCase))
                    filtered = filtered.Where(e => e.Name == PNs + "sp");
            }

            if (!string.IsNullOrWhiteSpace(selector.ShapeName))
            {
                var nameTarget = selector.ShapeName.Trim();
                filtered = filtered.Where(e =>
                {
                    var cNvPr = e.Descendants(PNs + "cNvPr").FirstOrDefault();
                    var name = cNvPr?.Attribute("name")?.Value;
                    return !string.IsNullOrWhiteSpace(name) && name.Contains(nameTarget, StringComparison.OrdinalIgnoreCase);
                });
            }

            if (!string.IsNullOrWhiteSpace(selector.TargetText))
            {
                var textTarget = NormalizePlainText(selector.TargetText);
                filtered = filtered.Where(e => NormalizePlainText(GetElementText(e)).Contains(textTarget, StringComparison.OrdinalIgnoreCase));
            }

            var list = filtered.ToList();
            if (selector.Ordinal.HasValue && selector.Ordinal.Value > 0 && selector.Ordinal.Value <= list.Count)
            {
                return new List<XElement> { list[selector.Ordinal.Value - 1] };
            }

            return list;
        }

        private static bool InchesEqual(double a, double b, double tolerance = 0.20)
        {
            return Math.Abs(a - b) <= tolerance;
        }

        private static double EmuToInches(long emu) => (double)emu / EmuPerInch;

        private static long GetShapeCoord(XElement shape, string axis)
        {
            var off = shape.Descendants(ANs + "off").FirstOrDefault();
            if (off != null)
            {
                var attr = off.Attribute(axis)?.Value;
                if (long.TryParse(attr, out var val))
                    return val;
            }
            return 0;
        }

        private static long GetShapeDimension(XElement shape, string dim)
        {
            var ext = shape.Descendants(ANs + "ext").FirstOrDefault();
            if (ext != null)
            {
                var attr = ext.Attribute(dim)?.Value;
                if (long.TryParse(attr, out var val))
                    return val;
            }
            return 0;
        }
    }
}
