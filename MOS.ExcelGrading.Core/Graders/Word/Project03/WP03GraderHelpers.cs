using System.Text.RegularExpressions;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project03
{
    internal static class WP03GraderHelpers
    {
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace Dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        public static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";

        private static readonly Regex WhiteSpaceRegex = new(@"\s+", RegexOptions.Compiled);

        public static void AddError(TaskResult result, string errorMessage, string fixAction)
        {
            result.Errors.Add(errorMessage);
            // Only add a single fixAction guidance per TaskResult to avoid
            // producing multiple, potentially duplicated suggestions.
            if (!string.IsNullOrWhiteSpace(fixAction)
                && result.FixActions.Count == 0)
            {
                result.FixActions.Add(fixAction);
            }
        }

        public static string NormalizeText(string? value)
        {
            return WhiteSpaceRegex.Replace((value ?? string.Empty).Trim(), " ");
        }

        public static IReadOnlyList<XElement> GetBodyElements(WordGradingContext context)
        {
            return context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Elements()
                .ToList()
                ?? new List<XElement>();
        }

        public static string GetParagraphText(XElement paragraph)
        {
            var text = string.Concat(paragraph.Descendants(W + "t").Select(node => node.Value));
            return NormalizeText(text);
        }

        public static string GetParagraphStyleId(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")
                ?.Element(W + "pStyle")
                ?.Attribute(W + "val")
                ?.Value
                ?? string.Empty;
        }

        public static bool IsHeadingParagraph(XElement paragraph)
        {
            var styleId = GetParagraphStyleId(paragraph);
            return styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsHeading1Paragraph(XElement paragraph)
        {
            return string.Equals(GetParagraphStyleId(paragraph), "Heading1", StringComparison.OrdinalIgnoreCase);
        }

        public static int FindParagraphIndexByExactText(IReadOnlyList<XElement> bodyElements, string expectedText)
        {
            var normalizedExpected = NormalizeText(expectedText);
            for (var i = 0; i < bodyElements.Count; i++)
            {
                if (bodyElements[i].Name != W + "p")
                {
                    continue;
                }

                var text = GetParagraphText(bodyElements[i]);
                if (string.Equals(text, normalizedExpected, StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }

        public static List<XElement> GetSectionElements(
            IReadOnlyList<XElement> bodyElements,
            int headingIndex,
            bool stopAtHeading1)
        {
            var elements = new List<XElement>();
            for (var i = headingIndex + 1; i < bodyElements.Count; i++)
            {
                var element = bodyElements[i];
                if (element.Name == W + "p")
                {
                    var shouldStop = stopAtHeading1
                        ? IsHeading1Paragraph(element)
                        : IsHeadingParagraph(element);
                    if (shouldStop)
                    {
                        break;
                    }
                }

                elements.Add(element);
            }

            return elements;
        }

        public static List<XElement> GetSectionParagraphs(
            IReadOnlyList<XElement> bodyElements,
            int headingIndex,
            bool stopAtHeading1)
        {
            return GetSectionElements(bodyElements, headingIndex, stopAtHeading1)
                .Where(element => element.Name == W + "p")
                .ToList();
        }

        public static XElement? FindParagraphContainingText(IEnumerable<XElement> paragraphs, string keyword)
        {
            var normalizedKeyword = NormalizeText(keyword);
            return paragraphs.FirstOrDefault(paragraph =>
                GetParagraphText(paragraph).Contains(normalizedKeyword, StringComparison.OrdinalIgnoreCase));
        }

        public static XElement? GetFirstTableAfterHeading(IReadOnlyList<XElement> bodyElements, int headingIndex)
        {
            for (var i = headingIndex + 1; i < bodyElements.Count; i++)
            {
                var element = bodyElements[i];
                if (element.Name == W + "tbl")
                {
                    return element;
                }

                if (element.Name == W + "p" && IsHeading1Paragraph(element))
                {
                    return null;
                }
            }

            return null;
        }

        public static string GetCoreTitle(WordGradingContext context)
        {
            return NormalizeText(context.CorePropertiesXml?.Root?.Element(Dc + "title")?.Value);
        }

        public static string ResolveWordPartEntry(string target)
        {
            var normalized = (target ?? string.Empty).Replace("\\", "/", StringComparison.Ordinal).Trim();
            while (normalized.StartsWith("../", StringComparison.Ordinal))
            {
                normalized = normalized[3..];
            }

            if (normalized.StartsWith("/", StringComparison.Ordinal))
            {
                normalized = normalized.TrimStart('/');
            }

            if (!normalized.StartsWith("word/", StringComparison.OrdinalIgnoreCase))
            {
                normalized = $"word/{normalized}";
            }

            return normalized;
        }

        public static bool TryGetRelatedXmlPart(
            WordGradingContext context,
            string relationshipId,
            out XDocument xmlPart,
            out string entryName)
        {
            xmlPart = null!;
            entryName = string.Empty;

            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                return false;
            }

            if (!context.TryGetDocumentRelationship(relationshipId, out var relationship))
            {
                return false;
            }

            entryName = ResolveWordPartEntry(relationship.Target);
            return context.TryGetXmlPart(entryName, out xmlPart);
        }

        public static bool TryGetDefaultHeaderPart(
            WordGradingContext context,
            out XDocument headerXml,
            out string headerEntryName)
        {
            headerXml = null!;
            headerEntryName = string.Empty;

            var sectPr = context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Element(W + "sectPr");

            if (sectPr == null)
            {
                return false;
            }

            var headerRef = sectPr.Elements(W + "headerReference")
                .FirstOrDefault(node => string.Equals(node.Attribute(W + "type")?.Value, "default", StringComparison.OrdinalIgnoreCase))
                ?? sectPr.Elements(W + "headerReference").FirstOrDefault();

            var relationshipId = headerRef?.Attribute(R + "id")?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                return false;
            }

            return TryGetRelatedXmlPart(context, relationshipId, out headerXml, out headerEntryName);
        }

        public static bool HasTitlePageEnabled(WordGradingContext context)
        {
            return context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Element(W + "sectPr")?
                .Element(W + "titlePg") != null;
        }

        public static string GetNumId(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")
                ?.Element(W + "numPr")
                ?.Element(W + "numId")
                ?.Attribute(W + "val")
                ?.Value
                ?? string.Empty;
        }

        public static string GetIlvl(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")
                ?.Element(W + "numPr")
                ?.Element(W + "ilvl")
                ?.Attribute(W + "val")
                ?.Value
                ?? string.Empty;
        }

        public static string GetSpacingLine(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")
                ?.Element(W + "spacing")
                ?.Attribute(W + "line")
                ?.Value
                ?? string.Empty;
        }

        public static string GetSpacingLineRule(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")
                ?.Element(W + "spacing")
                ?.Attribute(W + "lineRule")
                ?.Value
                ?? string.Empty;
        }
    }
}
