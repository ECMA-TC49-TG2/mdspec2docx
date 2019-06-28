using FSharp.Markdown;
using System.IO;

namespace MarkdownConverter.Spec
{
    internal class SourceLocation
    {
        public readonly string file;
        public readonly SectionRef section;
        public readonly MarkdownParagraph paragraph;
        public readonly MarkdownSpan span;
        public string _loc; // generated lazily. Of the form "file(line/col)" in a format recognizable by msbuild

        public SourceLocation(string file, SectionRef section, MarkdownParagraph paragraph, MarkdownSpan span)
        {
            this.file = file;
            this.section = section;
            this.paragraph = paragraph;
            this.span = span;
        }

        public string loc
        {
            get
            {
                // We have either a CurrentSection or a CurrentParagraph or both. Let's try to locate the error.
                // Ideally this would be present inside the FSharp.MarkdownParagraph or FSharp.MarkdownSpan class itself.
                // But it's not yet: https://github.com/tpetricek/FSharp.Formatting/issues/410
                // So for the time being, we have to hack around to try to find it.

                if (_loc != null) return _loc;

                if (file == null)
                {
                    _loc = "mdspec2docx";
                }
                else if (section == null && paragraph == null)
                {
                    _loc = file;
                }
                else
                {
                    var src = File.ReadAllText(file);

                    string src2 = src; int iOffset = 0;
                    bool foundSection = false, foundParagraph = false, foundSpan = false;
                    if (section != null)
                    {
                        var ss = Fuzzy.FindSection(src2, section);
                        if (ss != null) { src2 = src2.Substring(ss.start, ss.length); iOffset = ss.start; foundSection = true; }
                    }

                    if (paragraph != null)
                    {
                        var ss = Fuzzy.FindParagraph(src2, paragraph);
                        if (ss != null) { src2 = src2.Substring(ss.start, ss.length); iOffset += ss.start; foundParagraph = true; }
                        else
                        {
                            // If we can't find the paragraph within the current section, let's try to find it anywhere
                            ss = Fuzzy.FindParagraph(src, paragraph);
                            if (ss != null) { src2 = src.Substring(ss.start, ss.length); iOffset = ss.start; foundSection = false; foundParagraph = true; }
                        }
                    }

                    if (span != null)
                    {
                        var ss = Fuzzy.FindSpan(src2, span);
                        if (ss != null) { src2 = src.Substring(ss.start, ss.length); iOffset += ss.start; foundSpan = true; }
                    }

                    var startPos = iOffset;
                    var endPos = startPos + src2.Length;
                    int startLine, startCol, endLine, endCol;
                    if ((!foundSection && !foundParagraph) || !Fuzzy.FindLineCol(file, src, startPos, out startLine, out startCol, endPos, out endLine, out endCol))
                    {
                        _loc = file;
                    }
                    else
                    {
                        startLine += 1;
                        startCol += 1;
                        endLine += 1;
                        endCol += 1; // 1-based
                        if (foundSpan)
                        {
                            _loc = $"{file}({startLine},{startCol},{endLine},{endCol})";
                        }
                        else if (startLine == endLine)
                        {
                            _loc = $"{file}({startLine})";
                        }
                        else
                        {
                            _loc = $"{file}({startLine}-{endLine})";
                        }
                    }
                }
                return _loc;
            }
        }
    }

}
