using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project13
{
    internal static class WP13GraderHelpers
    {
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
        public static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
        public static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
        public static readonly XNamespace Dcterms = "http://purl.org/dc/terms/";
        public static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        public static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XNamespace Dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

        private static readonly Regex WhiteSpaceRegex = new(@"\s+", RegexOptions.Compiled);

        public static TaskResult CreateResult(string taskId, string taskName, decimal maxScore)
        {
            return new TaskResult
            {
                TaskId = taskId,
                TaskName = taskName,
                MaxScore = maxScore,
                Score = maxScore
            };
        }

        public static void AddError(TaskResult result, string errorMessage, string fixAction)
        {
            result.Score = 0m;
            result.Errors.Add(errorMessage);
            // Only add a single fixAction guidance per TaskResult to avoid
            // producing multiple, potentially duplicated suggestions.
            if (!string.IsNullOrWhiteSpace(fixAction)
                && result.FixActions.Count == 0)
            {
                result.FixActions.Add(fixAction);
            }
        }

        public static void AddDetail(TaskResult result, string detail)
        {
            if (!string.IsNullOrWhiteSpace(detail))
            {
                result.Details.Add(detail);
            }
        }

        public static string NormalizeText(string? value)
        {
            return WhiteSpaceRegex.Replace((value ?? string.Empty).Trim(), " ");
        }

        public static string GetParagraphText(XElement paragraph)
        {
            var text = string.Concat(paragraph
                .Descendants()
                .Where(node => node.Name == W + "t" || node.Name == W + "delText")
                .Select(node => node.Value));

            return NormalizeText(text);
        }

        public static string GetDocumentText(WordGradingContext context)
        {
            return NormalizeText(string.Join(" ", GetParagraphs(context).Select(GetParagraphText)));
        }

        public static IReadOnlyList<XElement> GetBodyElements(WordGradingContext context)
        {
            return context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Elements()
                .ToList()
                ?? new List<XElement>();
        }

        public static IReadOnlyList<XElement> GetParagraphs(WordGradingContext context)
        {
            return context.MainDocumentXml?
                .Descendants(W + "p")
                .ToList()
                ?? new List<XElement>();
        }

        public static IReadOnlyList<XElement> GetNonEmptyParagraphs(WordGradingContext context)
        {
            return GetParagraphs(context)
                .Where(p => !string.IsNullOrWhiteSpace(GetParagraphText(p)))
                .ToList();
        }

        public static bool HasExactlyLineSpacing(XElement paragraph, int expectedPoints)
        {
            var spacing = paragraph.Element(W + "pPr")?.Element(W + "spacing");
            var lineRule = spacing?.Attribute(W + "lineRule")?.Value ?? string.Empty;
            var line = spacing?.Attribute(W + "line")?.Value;

            if (!string.Equals(lineRule, "exact", StringComparison.OrdinalIgnoreCase)
                || !int.TryParse(line, NumberStyles.Integer, CultureInfo.InvariantCulture, out var lineValue))
            {
                return false;
            }

            return Math.Abs(lineValue - expectedPoints * 20) <= 2;
        }

        public static string GetParagraphStyle(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")?.Value ?? string.Empty;
        }

        public static bool HasStyle(XElement paragraph, string expectedStyle)
        {
            if (GetParagraphStyle(paragraph).Contains(expectedStyle, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return paragraph.Descendants(W + "rStyle")
                .Any(style => (style.Attribute(W + "val")?.Value ?? string.Empty)
                    .Contains(expectedStyle, StringComparison.OrdinalIgnoreCase));
        }

        public static bool HasImage(XElement paragraph)
        {
            return paragraph.Descendants(W + "drawing").Any()
                || paragraph.Descendants(W + "pict").Any();
        }

        public static int FindFirstImageParagraphIndex(IReadOnlyList<XElement> paragraphs)
        {
            for (var i = 0; i < paragraphs.Count; i++)
            {
                if (HasImage(paragraphs[i]))
                {
                    return i;
                }
            }

            return -1;
        }

        public static int FindProject13MainImageParagraphIndex(IReadOnlyList<XElement> paragraphs)
        {
            var imageParagraphIndexes = paragraphs
                .Select((paragraph, index) => new { paragraph, index })
                .Where(item => HasImage(item.paragraph))
                .Select(item => item.index)
                .ToList();

            foreach (var imageIndex in imageParagraphIndexes)
            {
                if (imageIndex < 4)
                {
                    continue;
                }

                var fourParagraphsAbove = paragraphs
                    .Skip(imageIndex - 4)
                    .Take(4)
                    .Select(GetParagraphText)
                    .ToList();

                var textAbove = NormalizeText(string.Join(" ", fourParagraphsAbove));
                var textAfter = NormalizeText(string.Join(" ", paragraphs
                    .Skip(imageIndex + 1)
                    .Take(5)
                    .Select(GetParagraphText)));

                var hasExpectedParagraphsAbove =
                    textAbove.Contains("This preview event will be hosted", StringComparison.OrdinalIgnoreCase)
                    && textAbove.Contains("She will give you practical advice", StringComparison.OrdinalIgnoreCase)
                    && textAbove.Contains("The event is aimed at people", StringComparison.OrdinalIgnoreCase)
                    && textAbove.Contains("We hope to have you join us", StringComparison.OrdinalIgnoreCase);

                var hasExpectedEventInfoAfter =
                    textAfter.Contains("This event will take place", StringComparison.OrdinalIgnoreCase)
                    || textAfter.Contains("margie@margiestravel.com", StringComparison.OrdinalIgnoreCase);

                if (hasExpectedParagraphsAbove && hasExpectedEventInfoAfter)
                {
                    return imageIndex;
                }
            }

            return FindContentImageParagraphIndex(paragraphs);
        }

        public static int FindContentImageParagraphIndex(IReadOnlyList<XElement> paragraphs)
        {
            var imageParagraphIndexes = paragraphs
                .Select((paragraph, index) => new { paragraph, index })
                .Where(item => HasImage(item.paragraph))
                .Select(item => item.index)
                .ToList();

            if (imageParagraphIndexes.Count == 0)
            {
                return -1;
            }

            var candidates = imageParagraphIndexes
                .Select(index => new
                {
                    Index = index,
                    NonEmptyBefore = paragraphs
                        .Take(index)
                        .Count(p => !string.IsNullOrWhiteSpace(GetParagraphText(p))),
                    NonEmptyAfter = paragraphs
                        .Skip(index + 1)
                        .Count(p => !string.IsNullOrWhiteSpace(GetParagraphText(p)))
                })
                .Where(candidate => candidate.NonEmptyBefore >= 4 && candidate.NonEmptyAfter >= 1)
                .ToList();

            if (candidates.Count > 0)
            {
                return candidates
                    .OrderByDescending(candidate => candidate.NonEmptyBefore)
                    .First()
                    .Index;
            }

            return imageParagraphIndexes.Last();
        }

        public static IReadOnlyList<XDocument> GetRelatedParts(
            WordGradingContext context,
            string relationshipTypeSuffix)
        {
            var parts = new List<XDocument>();

            foreach (var relationship in context.DocumentRelationships.Values)
            {
                if (!relationship.Type.EndsWith(relationshipTypeSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var entryName = ResolveWordPartEntry(relationship.Target);
                if (context.TryGetXmlPart(entryName, out var xml))
                {
                    parts.Add(xml);
                }
            }

            return parts;
        }

        public static string ResolveWordPartEntry(string target)
        {
            var normalized = (target ?? string.Empty).Replace("\\", "/", StringComparison.Ordinal).Trim();

            while (normalized.StartsWith("../", StringComparison.Ordinal))
            {
                normalized = normalized[3..];
            }

            normalized = normalized.TrimStart('/');

            if (!normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase))
            {
                normalized = $"word/{normalized}";
            }

            return normalized;
        }

        public static bool XmlPartHasVisibleText(XDocument part)
        {
            return part.Descendants(W + "t").Any(t => !string.IsNullOrWhiteSpace(t.Value))
                || part.Descendants(V + "shape").Any()
                || part.Descendants(V + "textbox").Any();
        }

        public static bool HasWatermark(XDocument part)
        {
            return part.Descendants(V + "shape").Any(shape =>
            {
                var style = shape.Attribute("style")?.Value ?? string.Empty;
                var type = shape.Attribute("type")?.Value ?? string.Empty;
                var id = shape.Attribute("id")?.Value ?? string.Empty;
                return style.Contains("rotation", StringComparison.OrdinalIgnoreCase)
                    || type.Contains("watermark", StringComparison.OrdinalIgnoreCase)
                    || id.Contains("PowerPlusWaterMarkObject", StringComparison.OrdinalIgnoreCase);
            })
            || part.Descendants().Any(node =>
                node.Name.LocalName.Contains("watermark", StringComparison.OrdinalIgnoreCase)
                || node.Value.Contains("watermark", StringComparison.OrdinalIgnoreCase));
        }

        public static DateTime? GetLastModifiedDate(WordGradingContext context)
        {
            var value = context.CorePropertiesXml?.Root?
                .Element(Dcterms + "modified")?
                .Value;

            return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var date)
                ? date.Date
                : null;
        }

        public static bool ContainsAnyFieldCode(XElement scope, params string[] expectedTokens)
        {
            var fieldText = string.Join(" ", scope.Descendants(W + "instrText").Select(node => node.Value));
            return expectedTokens.Any(token => fieldText.Contains(token, StringComparison.OrdinalIgnoreCase));
        }

        public static bool ContainsDateText(string text, DateTime expectedDate)
        {
            var normalizedText = NormalizeText(text);
            var cultures = new[]
            {
                CultureInfo.InvariantCulture,
                CultureInfo.GetCultureInfo("en-US"),
                CultureInfo.GetCultureInfo("en-GB"),
                CultureInfo.GetCultureInfo("vi-VN")
            };

            var formats = new[]
            {
                "M/d/yyyy",
                "MM/dd/yyyy",
                "M/d/yy",
                "MM/dd/yy",
                "d/M/yyyy",
                "dd/MM/yyyy",
                "d/M/yy",
                "dd/MM/yy",
                "yyyy-MM-dd",
                "dd-MM-yyyy",
                "d-M-yyyy",
                "dd-MMM-yyyy",
                "d-MMM-yyyy",
                "MMM d, yyyy",
                "MMMM d, yyyy",
                "d MMM yyyy",
                "d MMMM yyyy",
                "dddd, MMMM d, yyyy",
                "dddd, d MMMM yyyy"
            };

            foreach (var culture in cultures)
            {
                foreach (var format in formats)
                {
                    var candidate = expectedDate.ToString(format, culture);
                    if (normalizedText.Contains(candidate, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            var numericDateMatches = Regex.Matches(
                normalizedText,
                @"\b(?<a>\d{1,4})[\/\-.](?<b>\d{1,2})[\/\-.](?<c>\d{1,4})\b");

            foreach (Match match in numericDateMatches)
            {
                if (TryParseFlexibleNumericDate(match.Value, expectedDate))
                {
                    return true;
                }
            }

            return HasDatePartsNearEachOther(normalizedText, expectedDate);
        }

        private static bool TryParseFlexibleNumericDate(string value, DateTime expectedDate)
        {
            var formats = new[]
            {
                "M/d/yyyy",
                "MM/dd/yyyy",
                "M/d/yy",
                "MM/dd/yy",
                "d/M/yyyy",
                "dd/MM/yyyy",
                "d/M/yy",
                "dd/MM/yy",
                "yyyy/M/d",
                "yyyy/MM/dd",
                "yyyy-M-d",
                "yyyy-MM-dd",
                "d-M-yyyy",
                "dd-MM-yyyy",
                "d-M-yy",
                "dd-MM-yy"
            };

            return formats.Any(format =>
                DateTime.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
                && parsed.Date == expectedDate.Date);
        }

        private static bool HasDatePartsNearEachOther(string text, DateTime expectedDate)
        {
            var yearPattern = expectedDate.Year.ToString(CultureInfo.InvariantCulture);
            var dayPattern = expectedDate.Day.ToString(CultureInfo.InvariantCulture);
            var monthPattern = expectedDate.Month.ToString(CultureInfo.InvariantCulture);
            var englishMonthNames = new[]
            {
                expectedDate.ToString("MMM", CultureInfo.InvariantCulture),
                expectedDate.ToString("MMMM", CultureInfo.InvariantCulture)
            };
            var vietnameseMonthPattern = $"tháng {expectedDate.Month.ToString(CultureInfo.InvariantCulture)}";

            var dateLikeWindows = Regex.Matches(text, @".{0,40}" + Regex.Escape(yearPattern) + @".{0,40}");
            foreach (Match window in dateLikeWindows)
            {
                var value = window.Value;
                var hasDay = Regex.IsMatch(value, $@"\b0?{Regex.Escape(dayPattern)}\b");
                var hasNumericMonth = Regex.IsMatch(value, $@"\b0?{Regex.Escape(monthPattern)}\b");
                var hasNamedMonth = englishMonthNames.Any(month => value.Contains(month, StringComparison.OrdinalIgnoreCase))
                    || value.Contains(vietnameseMonthPattern, StringComparison.OrdinalIgnoreCase);

                if (hasDay && (hasNumericMonth || hasNamedMonth))
                {
                    return true;
                }
            }

            return false;
        }

        public static int? GetColumnsCountNearParagraph(XElement paragraph)
        {
            var cols = GetSectionPropertiesNearParagraph(paragraph)?.Element(W + "cols");
            if (cols == null)
            {
                return null;
            }

            return int.TryParse(cols.Attribute(W + "num")?.Value, out var count) ? count : 1;
        }

        public static int? GetColumnSpaceTwipsNearParagraph(XElement paragraph)
        {
            var cols = GetSectionPropertiesNearParagraph(paragraph)?.Element(W + "cols");
            return int.TryParse(cols?.Attribute(W + "space")?.Value, out var value) ? value : null;
        }

        public static IReadOnlyList<XElement> GetTables(WordGradingContext context)
        {
            return context.MainDocumentXml?
                .Descendants(W + "tbl")
                .ToList()
                ?? new List<XElement>();
        }

        public static bool HasTableTitleOrCaption(XElement table)
        {
            var tableProperties = table.Element(W + "tblPr");
            var caption = tableProperties?
                .Descendants(W + "tblCaption")
                .Select(node => node.Attribute(W + "val")?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            var description = tableProperties?
                .Descendants(W + "tblDescription")
                .Select(node => node.Attribute(W + "val")?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

            if (!string.IsNullOrWhiteSpace(caption) || !string.IsNullOrWhiteSpace(description))
            {
                return true;
            }

            return HasHeaderRowForAccessibility(table);
        }

        public static bool HasHeaderRowForAccessibility(XElement table)
        {
            var tableLook = table
                .Element(W + "tblPr")?
                .Element(W + "tblLook");

            var firstRowFlag = tableLook?.Attribute(W + "firstRow")?.Value;
            if (firstRowFlag == "1" || firstRowFlag?.Equals("true", StringComparison.OrdinalIgnoreCase) == true)
            {
                return true;
            }

            var firstRow = table.Elements(W + "tr").FirstOrDefault();
            if (firstRow == null)
            {
                return false;
            }

            var cnfStyle = firstRow
                .Element(W + "trPr")?
                .Element(W + "cnfStyle");

            return cnfStyle?.Attribute(W + "firstRow")?.Value == "1";
        }

        public static int FindBodyElementIndexContainingText(IReadOnlyList<XElement> bodyElements, string headingText)
        {
            for (var i = 0; i < bodyElements.Count; i++)
            {
                if (bodyElements[i].Name != W + "p")
                {
                    continue;
                }

                var text = GetParagraphText(bodyElements[i]);
                if (text.Contains(headingText, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        public static XElement? FindFirstTableAfterHeading(WordGradingContext context, string headingText)
        {
            var bodyElements = GetBodyElements(context);
            var headingIndex = FindBodyElementIndexContainingText(bodyElements, headingText);
            if (headingIndex < 0)
            {
                return null;
            }

            return bodyElements
                .Skip(headingIndex + 1)
                .TakeWhile(element => !IsLikelyHeadingParagraph(element))
                .FirstOrDefault(element => element.Name == W + "tbl");
        }

        public static IReadOnlyList<XElement> FindParagraphsAfterHeading(WordGradingContext context, string headingText)
        {
            var bodyElements = GetBodyElements(context);
            var headingIndex = FindBodyElementIndexContainingText(bodyElements, headingText);
            if (headingIndex < 0)
            {
                return new List<XElement>();
            }

            return bodyElements
                .Skip(headingIndex + 1)
                .TakeWhile(element => !IsLikelyHeadingParagraph(element))
                .Where(element => element.Name == W + "p")
                .ToList();
        }

        public static bool IsLikelyHeadingParagraph(XElement element)
        {
            if (element.Name != W + "p")
            {
                return false;
            }

            var style = GetParagraphStyle(element);
            if (style.Contains("Heading", StringComparison.OrdinalIgnoreCase)
                || style.Contains("Title", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var text = GetParagraphText(element);
            return text.Equals("Description", StringComparison.OrdinalIgnoreCase)
                || text.Equals("Filling Agents", StringComparison.OrdinalIgnoreCase)
                || text.Equals("Manufacturing Process", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsLandscapeSection(XElement sectionProperties)
        {
            var pageSize = sectionProperties.Element(W + "pgSz");
            if (pageSize == null)
            {
                return false;
            }

            var orient = pageSize.Attribute(W + "orient")?.Value ?? string.Empty;
            if (orient.Equals("landscape", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var widthValue = pageSize.Attribute(W + "w")?.Value;
            var heightValue = pageSize.Attribute(W + "h")?.Value;

            return int.TryParse(widthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var width)
                && int.TryParse(heightValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var height)
                && width > height;
        }

        public static IReadOnlyList<XElement> GetSectionPropertiesInDocumentOrder(WordGradingContext context)
        {
            return GetBodyElements(context)
                .Select(element => element.Name == W + "sectPr"
                    ? element
                    : element.Element(W + "pPr")?.Element(W + "sectPr"))
                .Where(sectPr => sectPr != null)
                .Cast<XElement>()
                .ToList();
        }

        public static IReadOnlyList<int> GetTableColumnWidthsTwips(XElement table)
        {
            var gridWidths = table
                .Element(W + "tblGrid")?
                .Elements(W + "gridCol")
                .Select(col => col.Attribute(W + "w")?.Value)
                .Where(value => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                .Select(value => int.Parse(value!, CultureInfo.InvariantCulture))
                .ToList();

            if (gridWidths != null && gridWidths.Count > 0)
            {
                return gridWidths;
            }

            return table
                .Descendants(W + "tr")
                .FirstOrDefault()?
                .Elements(W + "tc")
                .Select(cell => cell.Element(W + "tcPr")?.Element(W + "tcW")?.Attribute(W + "w")?.Value)
                .Where(value => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                .Select(value => int.Parse(value!, CultureInfo.InvariantCulture))
                .ToList()
                ?? new List<int>();
        }

        public static bool HasCitationPlaceholderAtEnd(XElement paragraph, string placeholderName)
        {
            var paragraphText = GetParagraphText(paragraph);
            if (!paragraphText.Contains(placeholderName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var fieldCode = string.Join(" ", paragraph.Descendants(W + "instrText").Select(node => node.Value));
            var sdtMetadata = string.Join(" ", paragraph
                .Descendants(W + "sdtPr")
                .Descendants()
                .Select(node => string.Join(" ", node.Attributes().Select(attribute => attribute.Value)) + " " + node.Value));

            var hasCitationArtifact = fieldCode.Contains("CITATION", StringComparison.OrdinalIgnoreCase)
                || fieldCode.Contains("BIBLIOGRAPHY", StringComparison.OrdinalIgnoreCase)
                || sdtMetadata.Contains("Citation", StringComparison.OrdinalIgnoreCase)
                || sdtMetadata.Contains("Bibliography", StringComparison.OrdinalIgnoreCase)
                || paragraph.ToString(SaveOptions.DisableFormatting).Contains("Citation", StringComparison.OrdinalIgnoreCase);

            var tailWindow = paragraphText.Length <= 80 ? paragraphText : paragraphText[^80..];
            return hasCitationArtifact && tailWindow.Contains(placeholderName, StringComparison.OrdinalIgnoreCase);
        }

        public static bool HasDrawingAltTextInScope(IEnumerable<XElement> scopeElements, string expectedAltText, bool requireSmartArt)
        {
            var scope = new XElement("scope", scopeElements);
            return scope.Descendants(W + "drawing").Any(drawing =>
            {
                if (requireSmartArt && !LooksLikeSmartArt(drawing))
                {
                    return false;
                }

                return drawing.Descendants(Wp + "docPr").Any(docPr =>
                {
                    var title = docPr.Attribute("title")?.Value ?? string.Empty;
                    var descr = docPr.Attribute("descr")?.Value ?? string.Empty;
                    var name = docPr.Attribute("name")?.Value ?? string.Empty;

                    return title.Equals(expectedAltText, StringComparison.OrdinalIgnoreCase)
                        || descr.Equals(expectedAltText, StringComparison.OrdinalIgnoreCase)
                        || name.Equals(expectedAltText, StringComparison.OrdinalIgnoreCase);
                });
            });
        }

        public static bool LooksLikeSmartArt(XElement drawing)
        {
            return drawing.Descendants(Dgm + "relIds").Any()
                || drawing.Descendants().Any(node =>
                    node.Name.NamespaceName.Contains("diagram", StringComparison.OrdinalIgnoreCase)
                    || node.Name.LocalName.Contains("diagram", StringComparison.OrdinalIgnoreCase)
                    || node.Name.LocalName.Contains("smartArt", StringComparison.OrdinalIgnoreCase));
        }

        public static bool HasInline3DModelInParagraph(XElement paragraph, WordGradingContext context, string? expectedName = null)
        {
            return paragraph.Descendants(Wp + "inline").Any(inline =>
            {
                var docPrText = string.Join(" ", inline
                    .Descendants(Wp + "docPr")
                    .Select(docPr => string.Join(" ", docPr.Attributes().Select(attribute => attribute.Value))));

                var inlineXml = inline.ToString(SaveOptions.DisableFormatting);
                var relationshipIds = inline
                    .Descendants()
                    .Attributes(R + "embed")
                    .Concat(inline.Descendants().Attributes(R + "link"))
                    .Concat(inline.Descendants().Attributes(R + "id"))
                    .Select(attribute => attribute.Value)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var relatedTargets = relationshipIds
                    .Where(context.DocumentRelationships.ContainsKey)
                    .Select(id => context.DocumentRelationships[id].Target)
                    .ToList();

                var modelSignals = string.Join(" ", docPrText, inlineXml, string.Join(" ", relatedTargets), string.Join(" ", context.Entries));
                var has3DModel = modelSignals.Contains("model3d", StringComparison.OrdinalIgnoreCase)
                    || modelSignals.Contains("3dmodel", StringComparison.OrdinalIgnoreCase)
                    || modelSignals.Contains("3D", StringComparison.OrdinalIgnoreCase);

                if (string.IsNullOrWhiteSpace(expectedName))
                {
                    return has3DModel;
                }

                var hasExpectedName = modelSignals.Contains(expectedName, StringComparison.OrdinalIgnoreCase)
                    || modelSignals.Contains(expectedName.Replace(" ", string.Empty, StringComparison.Ordinal), StringComparison.OrdinalIgnoreCase)
                    || modelSignals.Contains(expectedName.Replace(" ", "-", StringComparison.Ordinal), StringComparison.OrdinalIgnoreCase)
                    || modelSignals.Contains(expectedName.Replace(" ", "_", StringComparison.Ordinal), StringComparison.OrdinalIgnoreCase);

                return has3DModel && hasExpectedName;
            });
        }

        public static bool HasAnchored3DModelInScope(IEnumerable<XElement> scopeElements)
        {
            var scope = new XElement("scope", scopeElements);
            return scope.Descendants(Wp + "anchor").Any(anchor =>
            {
                var xml = anchor.ToString(SaveOptions.DisableFormatting);
                return xml.Contains("model3d", StringComparison.OrdinalIgnoreCase)
                    || xml.Contains("3dmodel", StringComparison.OrdinalIgnoreCase)
                    || xml.Contains("Blister", StringComparison.OrdinalIgnoreCase);
            });
        }

        private static XElement? GetSectionPropertiesNearParagraph(XElement paragraph)
        {
            var ownSectionProperties = paragraph.Element(W + "pPr")?.Element(W + "sectPr");
            if (ownSectionProperties != null)
            {
                return ownSectionProperties;
            }

            var nextParagraphSectionProperties = paragraph
                .ElementsAfterSelf(W + "p")
                .Select(p => p.Element(W + "pPr")?.Element(W + "sectPr"))
                .FirstOrDefault(sectPr => sectPr != null);

            if (nextParagraphSectionProperties != null)
            {
                return nextParagraphSectionProperties;
            }

            var directFollowingSectionProperties = paragraph
                .ElementsAfterSelf(W + "sectPr")
                .FirstOrDefault();

            if (directFollowingSectionProperties != null)
            {
                return directFollowingSectionProperties;
            }

            return paragraph.Ancestors().LastOrDefault()?.Element(W + "body")?.Element(W + "sectPr");
        }
    }
}

