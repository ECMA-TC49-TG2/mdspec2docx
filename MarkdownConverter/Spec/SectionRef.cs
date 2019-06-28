using FSharp.Markdown;
using System;
using System.Linq;

namespace MarkdownConverter.Spec
{
    internal class SectionRef
    {
        public string Number;        // "10.1.2"
        public string Title;         // "Goto Statement"
        public int Level;            // 1-based level, e.g. 3
        public string Url;           // statements.md#goto-statement
        public string BookmarkName;  // _Toc00023
        public string MarkdownTitle; // "Goto Statement" or "`<code>`"
        public SourceLocation Loc;
        public static int count = 1;

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
