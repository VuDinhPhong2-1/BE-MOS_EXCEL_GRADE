using System.Text.RegularExpressions;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project05
{
    internal static class WP05GraderHelpers
    {
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        private static readonly Regex WhiteSpaceRegex = new(@"\s+", RegexOptions.Compiled);

        public static string NormalizeText(string? value)
        {
            return WhiteSpaceRegex.Replace((value ?? string.Empty).Trim(), " ");
        }

        public static string GetParagraphText(XElement paragraph)
        {
            var text = string.Concat(paragraph.Descendants(W + "t").Select(node => node.Value));
            return NormalizeText(text);
        }

        public static IReadOnlyList<XElement> GetBodyElements(WordGradingContext context)
        {
            return context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Elements()
                .ToList()
                ?? new List<XElement>();
        }

        public static IReadOnlyList<XElement> GetBodyParagraphs(WordGradingContext context)
        {
            return context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Descendants(W + "p")
                .ToList()
                ?? new List<XElement>();
        }

        public static string GetDocumentText(WordGradingContext context)
        {
            var text = string.Concat(
                context.MainDocumentXml?.Descendants(W + "t").Select(node => node.Value) ?? Enumerable.Empty<string>());
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

        public static string GetTableCellText(XElement tableCell)
        {
            var text = string.Concat(tableCell.Descendants(W + "t").Select(node => node.Value));
            return NormalizeText(text);
        }

        public static bool HasTabCharacter(XElement paragraph)
        {
            if (paragraph.Descendants(W + "tab").Any())
            {
                return true;
            }

            var rawText = string.Concat(paragraph.Descendants(W + "t").Select(node => node.Value));
            return rawText.Contains('\t', StringComparison.Ordinal);
        }

        public static bool HasExactWord(string text, string word)
        {
            if (string.IsNullOrWhiteSpace(word))
            {
                return false;
            }

            return Regex.IsMatch(
                text ?? string.Empty,
                $@"(?<!\w){Regex.Escape(word)}(?!\w)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        public static int CountExactPhrase(string text, string phrase, bool ignoreCase)
        {
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(phrase))
            {
                return 0;
            }

            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            var count = 0;
            var start = 0;

            while (start <= text.Length - phrase.Length)
            {
                var index = text.IndexOf(phrase, start, comparison);
                if (index < 0)
                {
                    break;
                }

                count++;
                start = index + phrase.Length;
            }

            return count;
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

        public static List<XElement> GetTocSdtNodes(WordGradingContext context)
        {
            return context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Elements(W + "sdt")
                .Where(node =>
                    node.Element(W + "sdtPr")?
                        .Element(W + "docPartObj")?
                        .Element(W + "docPartGallery")?
                        .Attribute(W + "val")?
                        .Value
                        .Equals("Table of Contents", StringComparison.OrdinalIgnoreCase) == true)
                .ToList()
                ?? new List<XElement>();
        }

        public static string GetTextboxText(XElement textboxNode)
        {
            var text = string.Concat(textboxNode.Descendants(W + "t").Select(node => node.Value));
            return NormalizeText(text);
        }

        public static string? TryGetHyperlinkIdFromDrawing(XElement drawingNode)
        {
            return drawingNode.Descendants(A + "hlinkClick")
                .Select(node => node.Attribute(R + "id")?.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        }

        public static bool HasBookmark(WordGradingContext context, string bookmarkName)
        {
            if (string.IsNullOrWhiteSpace(bookmarkName))
            {
                return false;
            }

            return context.MainDocumentXml?.Descendants(W + "bookmarkStart")
                .Any(node => string.Equals(
                    node.Attribute(W + "name")?.Value,
                    bookmarkName,
                    StringComparison.Ordinal)) == true;
        }
    }
}

