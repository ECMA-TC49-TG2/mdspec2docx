using FSharp.Markdown;
using System;
using System.Linq;

namespace MarkdownConverter.Spec
{
    internal class SectionRef
    {
        // TODO: It would be nice if this could be read-only, but it's computed based on Level.
        // (We could create a new SectionRef with the new number, admittedly, but it's probably not worth it.)
        /// <summary>
        /// Section number, e.g. 10.1.2
        /// </summary>
        public string Number { get; set; }

        /// <summary>
        /// Section title, e.g. "Goto Statement"
        /// </summary>
        public string Title { get; }

        /// <summary>
        /// 1-based level, e.g. 3
        /// </summary>
        public int Level { get; }

        /// <summary>
        /// URL for the Markdown source, e.g. statements.md#goto-statement
        /// </summary>
        public string Url { get; }

        /// <summary>
        /// Name of generated bookmark, e.g. _Toc00023.
        /// </summary>
        public string BookmarkName { get; }

        /// <summary>
        /// Title in Markdown, e.g. "Goto Statement" or "`<code>`"
        /// </summary>
        public string MarkdownTitle; // "Goto Statement" or "`<code>`"

        /// <summary>
        /// Location in source Markdown.
        /// </summary>
        public SourceLocation Loc { get; }

        /// <summary>
        /// Counter used to generate bookmarks.
        /// </summary>
        private static int count = 1;

        public SectionRef(MarkdownParagraph.Heading mdh, string filename)
        {
            Level = mdh.size;
            var spans = mdh.body;
            if (spans.Length == 1 && spans.First().IsLiteral)
            {
                Title = MarkdownUtilities.UnescapeLiteral(spans.First() as MarkdownSpan.Literal).Trim();
                MarkdownTitle = Title;
            }
            else if (spans.Length == 1 && spans.First().IsInlineCode)
            {
                Title = (spans.First() as MarkdownSpan.InlineCode).code.Trim();
                MarkdownTitle = "`" + Title + "`";
            }
            else
            {
                throw new NotSupportedException("Heading must be a single literal/inlinecode");
            }
            foreach (var c in Title)
            {
                if (c >= 'a' && c <= 'z') Url += c;
                else if (c >= 'A' && c <= 'Z') Url += char.ToLowerInvariant(c);
                else if (c >= '0' && c <= '9') Url += c;
                else if (c == '-' || c == '_') Url += c;
                else if (c == ' ') Url += '-';
            }
            Url = filename + "#" + Url;
            BookmarkName = $"_Toc{count:00000}"; count++;
            Loc = new SourceLocation(filename, this, mdh, null);
        }

        public override string ToString() => Url;
    }
}
