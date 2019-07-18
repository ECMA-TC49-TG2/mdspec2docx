using FSharp.Formatting.Common;
using FSharp.Markdown;

namespace MarkdownConverter.Spec
{
    internal class SourceLocation
    {
        public string File { get; }
        public SectionRef Section { get; }
        public MarkdownParagraph Paragraph { get; }
        public MarkdownSpan Span { get; }
        public string _loc; // generated lazily.

        public SourceLocation(string file, SectionRef section, MarkdownParagraph paragraph, MarkdownSpan span)
        {
            File = file;
            Section = section;
            Paragraph = paragraph;
            Span = span;
        }

        /// <summary>
        /// Description of the location, of the form "file(line/col)" in a format recognizable by msbuild.
        /// </summary>
        public string Description
        {
            get
            {
                // We have either a CurrentSection or a CurrentParagraph or both. Let's try to locate the error.
                // Ideally this would be present inside the FSharp.MarkdownParagraph or FSharp.MarkdownSpan class itself.
                // But it's not yet: https://github.com/tpetricek/FSharp.Formatting/issues/410
                // So for the time being, we have to hack around to try to find it.

                if (_loc != null)
                {
                    return _loc;
                }

                if (File == null)
                {
                    _loc = "mdspec2docx";
                }
                else if (Section == null && Paragraph == null)
                {
                    _loc = File;
                }
                else
                {

                    var src = System.IO.File.ReadAllText(File);

                    string src2 = src; int iOffset = 0;
                    bool foundSection = false, foundParagraph = false, foundSpan = false;
                    if (Section != null)
                    {
                        var ss = Fuzzy.FindSection(src2, Section);
                        if (ss != null) { src2 = src2.Substring(ss.Start, ss.Length); iOffset = ss.Start; foundSection = true; }
                    }

                    if (Paragraph != null)
                    {
                        var ss = Fuzzy.FindParagraph(src2, Paragraph);
                        if (ss != null) { src2 = src2.Substring(ss.Start, ss.Length); iOffset += ss.Start; foundParagraph = true; }
                        else
                        {
                            // If we can't find the paragraph within the current section, let's try to find it anywhere
                            ss = Fuzzy.FindParagraph(src, Paragraph);
                            if (ss != null) { src2 = src.Substring(ss.Start, ss.Length); iOffset = ss.Start; foundSection = false; foundParagraph = true; }
                        }
                    }

                    if (Span != null)
                    {
                        var ss = Fuzzy.FindSpan(src2, Span);
                        if (ss != null) { src2 = src.Substring(ss.Start, ss.Length); iOffset += ss.Start; foundSpan = true; }
                    }

                    var startPos = iOffset;
                    var endPos = startPos + src2.Length;
                    int startLine, startCol, endLine, endCol;
                    if ((!foundSection && !foundParagraph) || !Fuzzy.FindLineCol(File, src, startPos, out startLine, out startCol, endPos, out endLine, out endCol))
                    {
                        _loc = File;
                    }
                    else
                    {
                        startLine += 1;
                        startCol += 1;
                        endLine += 1;
                        endCol += 1; // 1-based
                        if (foundSpan)
                        {
                            _loc = $"{File}({startLine},{startCol},{endLine},{endCol})";
                        }
                        else if (startLine == endLine)
                        {
                            _loc = $"{File}({startLine})";
                        }
                        else
                        {
                            _loc = $"{File}({startLine}-{endLine})";
                        }
                    }
                }
                return _loc;
            }
        }
    }

}
