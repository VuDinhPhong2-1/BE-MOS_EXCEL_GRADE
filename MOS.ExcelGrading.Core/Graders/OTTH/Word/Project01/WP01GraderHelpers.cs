using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.OTTH.Word.Project01
{
    internal static class WP01GraderHelpers
    {
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
        public static readonly XNamespace AM3D = "http://schemas.microsoft.com/office/drawing/2017/model3d";
        public static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

        private static readonly Regex WhiteSpaceRegex = new(@"\s+", RegexOptions.Compiled);

        public static string NormalizeText(string? value)
        {
            return WhiteSpaceRegex.Replace((value ?? string.Empty).Trim(), " ");
        }

        public static string NormalizeSortText(string? value)
        {
            var normalized = RemoveDiacritics((value ?? string.Empty).Trim());
            normalized = WhiteSpaceRegex.Replace(normalized, " ");
            return normalized.ToUpperInvariant();
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
            var styleId = GetParagraphStyleId(paragraph);
            return string.Equals(styleId, "Heading1", StringComparison.OrdinalIgnoreCase);
        }

        public static IReadOnlyList<XElement> GetBodyElements(WordGradingContext context)
        {
            return context.MainDocumentXml?.Root?
                .Element(W + "body")?
                .Elements()
                .ToList()
                ?? new List<XElement>();
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

        public static List<XElement> GetSectionParagraphs(
            IReadOnlyList<XElement> bodyElements,
            int headingIndex,
            bool stopAtHeading1)
        {
            var paragraphs = new List<XElement>();
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

                    var text = GetParagraphText(element);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        paragraphs.Add(element);
                    }
                }
            }

            return paragraphs;
        }

        public static XElement? GetFirstTableAfterHeading(
            IReadOnlyList<XElement> bodyElements,
            int headingIndex)
        {
            for (var i = headingIndex + 1; i < bodyElements.Count; i++)
            {
                var element = bodyElements[i];
                if (element.Name == W + "tbl")
                {
                    return element;
                }

                if (element.Name == W + "p" && IsHeadingParagraph(element))
                {
                    return null;
                }
            }

            return null;
        }

        public static XElement? FindParagraphContainingText(IEnumerable<XElement> paragraphs, string keyword)
        {
            var normalizedKeyword = NormalizeText(keyword);
            return paragraphs.FirstOrDefault(paragraph =>
                GetParagraphText(paragraph).Contains(normalizedKeyword, StringComparison.OrdinalIgnoreCase));
        }

        public static string CanonicalXml(XElement? element)
        {
            if (element == null)
            {
                return string.Empty;
            }

            var clone = new XElement(element);
            RemoveVolatileAttributes(clone);
            return clone.ToString(SaveOptions.DisableFormatting);
        }

        public static int ParseGeologicPeriodOrder(string value)
        {
            var text = NormalizeText(value);
            if (string.IsNullOrWhiteSpace(text))
            {
                return int.MaxValue;
            }

            var match = Regex.Match(text, @"^(?<rank>\d+)\s*\.?\s*(?<name>.+)$");
            if (match.Success && int.TryParse(match.Groups["rank"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rank))
            {
                return rank;
            }

            if (text.Contains("TRIASSIC", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            if (text.Contains("JURASSIC", StringComparison.OrdinalIgnoreCase))
            {
                return 2;
            }

            if (text.Contains("CRETACEOUS", StringComparison.OrdinalIgnoreCase))
            {
                return 3;
            }

            if (text.Contains("CENOZOIC", StringComparison.OrdinalIgnoreCase))
            {
                return 4;
            }

            return int.MaxValue;
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

        private static void RemoveVolatileAttributes(XElement element)
        {
            foreach (var descendant in element.DescendantsAndSelf())
            {
                var volatileAttributes = descendant.Attributes()
                    .Where(attribute =>
                        attribute.Name.LocalName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase)
                        || attribute.Name.LocalName.Equals("paraId", StringComparison.OrdinalIgnoreCase)
                        || attribute.Name.LocalName.Equals("textId", StringComparison.OrdinalIgnoreCase)
                        || attribute.Name.LocalName.Equals("anchorId", StringComparison.OrdinalIgnoreCase)
                        || attribute.Name.LocalName.Equals("editId", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var attribute in volatileAttributes)
                {
                    attribute.Remove();
                }
            }
        }

        private static string RemoveDiacritics(string value)
        {
            var normalized = value.Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder(normalized.Length);
            foreach (var character in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(character);
                }
            }

            return builder.ToString().Normalize(NormalizationForm.FormC);
        }
    }
}

