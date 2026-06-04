using System.Text.RegularExpressions;
using System.Xml.Linq;
using MOS.ExcelGrading.Core.Models;

namespace MOS.ExcelGrading.Core.Graders.Word.Project09
{
    internal static class WP09GraderHelpers
    {
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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

            if (!string.IsNullOrWhiteSpace(fixAction)
                && !result.FixActions.Contains(fixAction, StringComparer.Ordinal))
            {
                result.FixActions.Add(fixAction);
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

        public static int FindParagraphIndexContaining(IReadOnlyList<XElement> bodyElements, string expectedText)
        {
            var normalizedExpected = NormalizeText(expectedText);
            for (var i = 0; i < bodyElements.Count; i++)
            {
                if (bodyElements[i].Name != W + "p")
                {
                    continue;
                }

                if (GetParagraphText(bodyElements[i]).Contains(normalizedExpected, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        public static XElement? GetFirstTableAfter(IReadOnlyList<XElement> bodyElements, int startIndex)
        {
            for (var i = Math.Max(0, startIndex + 1); i < bodyElements.Count; i++)
            {
                var element = bodyElements[i];
                if (element.Name == W + "tbl")
                {
                    return element;
                }

                if (element.Name == W + "p" && i > startIndex + 1)
                {
                    var text = GetParagraphText(element);
                    var style = element.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")?.Value ?? string.Empty;
                    if (style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                        || text.Equals("Preserve your greatest memories!", StringComparison.OrdinalIgnoreCase))
                    {
                        return null;
                    }
                }
            }

            return null;
        }

        public static bool HasParagraphNumbering(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId") != null;
        }

        public static bool HasUnresolvedTrackedChanges(WordGradingContext context)
        {
            var document = context.MainDocumentXml;
            if (document?.Root == null)
            {
                return false;
            }

            var revisionNames = new HashSet<XName>
            {
                W + "ins",
                W + "del",
                W + "moveFrom",
                W + "moveTo",
                W + "rPrChange",
                W + "pPrChange",
                W + "tblPrChange",
                W + "tcPrChange",
                W + "trPrChange",
                W + "sectPrChange"
            };

            return document.Descendants().Any(node => revisionNames.Contains(node.Name));
        }

        public static XDocument? GetFootnotesPart(WordGradingContext context)
        {
            if (context.TryGetXmlPart("word/footnotes.xml", out var footnotesXml))
            {
                return footnotesXml;
            }

            var footnoteRelationship = context.DocumentRelationships.Values.FirstOrDefault(rel =>
                rel.Type.EndsWith("/footnotes", StringComparison.OrdinalIgnoreCase));

            if (footnoteRelationship == null)
            {
                return null;
            }

            var target = ResolveWordPartEntry(footnoteRelationship.Target);
            return context.TryGetXmlPart(target, out footnotesXml) ? footnotesXml : null;
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
    }
}