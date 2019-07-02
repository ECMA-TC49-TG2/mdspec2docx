using FSharp.Markdown;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

/// <summary>
/// FIXME: This is horrible...
/// </summary>
namespace MarkdownConverter.Spec
{
    internal static class Fuzzy
    {
        private static readonly Dictionary<string, List<int>> linestarts = new Dictionary<string, List<int>>();

        public static bool FindLineCol(string fn, string src, int startPos, out int startLine, out int startCol, int endPos, out int endLine, out int endCol)
        {
            startLine = -1; startCol = -1; endLine = -1; endCol = -1;

            List<int> starts;
            if (!linestarts.TryGetValue(fn, out starts))
            {
                starts = FindRawLines(src).Select(raw => raw.span.Start).ToList();
                linestarts[fn] = starts;
            }

            for (int i = 0; i < starts.Count; i++)
            {
                if (starts[i] <= startPos && (i == starts.Count || starts[i + 1] > startPos))
                {
                    startLine = i;
                    startCol = startPos - starts[i];
                }
                if (starts[i] <= endPos && (i == starts.Count || starts[i + 1] > endPos))
                {
                    endLine = i;
                    endCol = endPos - starts[i];
                }
            }
            if (startLine == -1 || startCol == -1 || endLine == -1 || endCol == -1) return false;
            return true;
        }

        public static Span FindSpan(string src, MarkdownSpan mds)
        {
            var mdwords = Span2Words(mds).ToList();
            var srcwords = FindWords(src).ToList();
            var srcwordsProjection = srcwords.Select(w => w.word);
            var s = LevenshteinSearch(mdwords, srcwordsProjection);
            if (s == null) return null;
            var wordStart = s.Start;
            var wordEnd = s.Start + s.Length;
            var spanStart = srcwords[wordStart].span.Start;
            var spanEnd = (wordEnd == srcwords.Count ? srcwords.Last().span.Start + srcwords.Last().span.Length : srcwords[wordEnd].span.Start);
            return new Span(spanStart, spanEnd - spanStart);
        }

        private static bool IsWordChar(char c)
        {
            return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '\'';
        }

        private static IEnumerable<WordSpan> FindWords(string src)
        {
            int i = 0;
            while (i < src.Length && !IsWordChar(src[i])) i++;
            while (i < src.Length)
            {
                string word = "";
                int spanStart = i;
                while (i < src.Length && IsWordChar(src[i])) { word += src[i]; i++; }
                yield return new WordSpan { word = word, span = new Span(spanStart, i - spanStart) };
                while (i < src.Length && !IsWordChar(src[i])) i++;
            }
        }

        public static Span FindParagraph(string src, MarkdownParagraph mdp)
        {
            var mdwords = Paragraph2Words(mdp).ToList();
            var mdwords2 = string.Join(",", mdwords);
            var paragraphs = from paragraph in FindParagraphs(src)
                             let pwords = String2Words(paragraph.text)
                             let i = LevenshteinDistance(mdwords, pwords)
                             orderby i ascending
                             select paragraph.span;

            return paragraphs.FirstOrDefault();
        }

        private class WordSpan
        {
            public string word;
            public Span span;
        }

        private class TextSpan
        {
            public string text;
            public Span span;
        }

        private class LineSpan
        {
            public string line;
            public Span span;
        }

        private class LineOrCodeblockSpan
        {
            public string line;
            public string codeblock;
            public Span span;
        }

        private static IEnumerable<TextSpan> FindParagraphs(string src)
        {
            var lines = FindLinesAndCodeblocks(src).ToList();
            int i = 0;
            while (i < lines.Count && lines[i].line != null && lines[i].line == "") i++;
            while (i < lines.Count)
            {
                if (lines[i].codeblock != null)
                {
                    yield return new TextSpan { text = lines[i].codeblock, span = lines[i].span };
                    i++;
                }
                else
                {
                    var p = "";
                    var spanStart = lines[i].span.Start;
                    var lastLine = lines[i];
                    while (i < lines.Count && lines[i].line != null && lines[i].line != "") { lastLine = lines[i]; p += lines[i].line + "\r\n"; i++; }
                    yield return new TextSpan { text = p, span = new Span(spanStart, lastLine.span.End - spanStart) };
                }
                while (i < lines.Count && lines[i].line != null && lines[i].line == "") i++;
            }
        }

        private static IEnumerable<LineOrCodeblockSpan> FindLinesAndCodeblocks(string src)
        {
            var lines = FindRawLines(src).ToList();
            string lang = null, fence = null, indent = null, terminator = "";
            StringBuilder fb = null, cb = null; int cbSpanStart = 0;
            foreach (var line in lines)
            {
                if (cb == null) // non-codeblock
                {
                    MdIsFenceStart(line.line, out lang, out fence, out indent, out terminator);
                    if (lang == null) yield return new LineOrCodeblockSpan { line = line.line, span = line.span };
                    else { fb = new StringBuilder(); fb.Append(line); cb = new StringBuilder(); cbSpanStart = line.span.Start; }
                }
                else // codeblock
                {
                    fb.Append(line);
                    var line2 = MdRemoveFenceIndent(line.line.TrimEnd("\r\n".ToCharArray()), indent);
                    if (!MdIsFenceEnd(line2, fence)) { cb.AppendLine(line2); continue; }
                    var code = cb.ToString();
                    yield return new LineOrCodeblockSpan { codeblock = code, span = new Span(cbSpanStart, line.span.Start + line.span.Length - cbSpanStart) };
                    cb = null;
                }
            }
        }

        private static Regex MdFenceStart = new Regex("^( *)((?<back>````*)|(?<tilde>~~~~*)) *([^ \r\n]*)");

        private static bool MdIsFenceStart(string line, out string lang, out string fence, out string indent, out string terminator)
        {
            lang = null; fence = null; indent = null; terminator = null;
            var m = MdFenceStart.Match(line);
            if (!m.Success) return false;
            indent = m.Groups[1].Value;
            fence = m.Groups[2].Value;
            lang = m.Groups[3].Value;
            if (line.EndsWith("\r\n")) terminator = "\r\n";
            else if (line.EndsWith("\r")) terminator = "\r";
            else if (line.EndsWith("\n")) terminator = "\n";
            else terminator = "\r\n";
            return true;
        }

        private static string MdRemoveFenceIndent(string line, string indent)
        {
            while (indent.StartsWith(" ") && line.StartsWith(" "))
            {
                line = line.Substring(1); indent = indent.Substring(1);
            }
            return line;
        }

        private static bool MdIsFenceEnd(string line, string fence)
        {
            if (!line.StartsWith(fence)) return false;
            fence = fence.Substring(0, 1);
            while (line.StartsWith(fence)) line = line.Substring(1);
            while (line.StartsWith(" ")) line = line.Substring(1);
            return (line == "");
        }

        private static IEnumerable<LineSpan> FindRawLines(string src)
        {
            for (int i = 0; i < src.Length;)
            {
                var lineSpanStart = i;
                var line = "";
                while (i < src.Length && src[i] != '\r' && src[i] != '\n') { line += src[i]; i++; }
                var lineSpanLength = i - lineSpanStart;
                if (i == src.Length) i += 0;
                else if (i == src.Length - 1) i += 1;
                else if (i <= src.Length - 2 && src[i] == '\r' && src[i + 1] == '\n') i += 2;
                else if (i <= src.Length - 2 && src[i] == '\r' && src[i + 1] != '\n') i += 1;
                else if (i <= src.Length - 2 && src[i] == '\n' && src[i + 1] == '\r') i += 2;
                else if (i <= src.Length - 2 && src[i] == '\n' && src[i + 1] != '\r') i += 1;
                yield return new LineSpan { line = line, span = new Span(lineSpanStart, lineSpanLength) };
            }
        }

        private static IEnumerable<string> Spans2Words(IEnumerable<MarkdownSpan> spans)
        {
            foreach (var mds in spans) foreach (var s in Span2Words(mds)) yield return s;
        }

        private static IEnumerable<string> Span2Words(MarkdownSpan md)
        {
            if (md.IsLiteral)
            {
                var mdl = md as MarkdownSpan.Literal;
                foreach (var s in String2Words(MarkdownUtilities.UnescapeLiteral(mdl))) yield return s;
            }
            else if (md.IsStrong)
            {
                var mds = md as MarkdownSpan.Strong;
                foreach (var s in Spans2Words(mds.body)) yield return s;
            }
            else if (md.IsEmphasis)
            {
                var mde = md as MarkdownSpan.Emphasis;
                foreach (var s in Spans2Words(mde.body)) yield return s;
            }
            else if (md.IsInlineCode)
            {
                var mdi = md as MarkdownSpan.InlineCode;
                foreach (var s in String2Words(mdi.code)) yield return s;
            }
            else if (md.IsDirectLink)
            {
                var mdl = md as MarkdownSpan.DirectLink;
                foreach (var s in Spans2Words(mdl.body)) yield return s;
                foreach (var s in String2Words(mdl.title.Option())) yield return s;
            }
            else if (md.IsIndirectLink)
            {
                var mdl = md as MarkdownSpan.DirectLink;
                foreach (var s in Spans2Words(mdl.body)) yield return s;
                foreach (var s in String2Words(mdl.link)) yield return s;
                foreach (var s in String2Words(mdl.title.Option())) yield return s;
            }
            else if (md.IsAnchorLink)
            {
                var mdl = md as MarkdownSpan.AnchorLink;
                foreach (var s in String2Words(mdl.link)) yield return s;
            }
            else if (md.IsDirectImage)
            {
                var mdi = md as MarkdownSpan.DirectImage;
                foreach (var s in String2Words(mdi.body)) yield return s;
                foreach (var s in String2Words(mdi.link)) yield return s;
                foreach (var s in String2Words(mdi.title.Option())) yield return s;
            }
            else if (md.IsIndirectImage)
            {
                var mdi = md as MarkdownSpan.IndirectImage;
                foreach (var s in String2Words(mdi.body)) yield return s;
                foreach (var s in String2Words(mdi.link)) yield return s;
                foreach (var s in String2Words(mdi.key)) yield return s;
            }
            else if (md.IsEmbedSpans)
            {
            }
            else if (md.IsHardLineBreak)
            {
            }
            else if (md.IsLatexDisplayMath)
            {
                var mdl = md as MarkdownSpan.LatexDisplayMath;
                foreach (var s in String2Words(mdl.code)) yield return s;
            }
            else if (md.IsLatexInlineMath)
            {
                var mdl = md as MarkdownSpan.LatexInlineMath;
                foreach (var s in String2Words(mdl.code)) yield return s;
            }
        }

        private static IEnumerable<string> Paragraphs2Words(IEnumerable<MarkdownParagraph> ps)
        {
            foreach (var p in ps) foreach (var s in Paragraph2Words(p)) yield return s;
        }

        private static IEnumerable<string> Paragraph2Words(MarkdownParagraph md)
        {
            if (md.IsHeading)
            {
                var mdh = md as MarkdownParagraph.Heading;
                foreach (var s in Spans2Words(mdh.body)) yield return s;
            }
            else if (md.IsParagraph)
            {
                var mdp = md as MarkdownParagraph.Paragraph;
                foreach (var s in Spans2Words(mdp.body)) yield return s;
            }
            else if (md.IsListBlock)
            {
                var mdl = md as MarkdownParagraph.ListBlock;
                foreach (var bullet in mdl.items) foreach (var s in Paragraphs2Words(bullet)) yield return s;
            }
            else if (md.IsCodeBlock)
            {
                var mdc = md as MarkdownParagraph.CodeBlock;
                var code = mdc.code;
                var lang = mdc.language?.Trim();
                if (!string.IsNullOrWhiteSpace(lang)) yield return lang;
                foreach (var s in String2Words(code)) yield return s;
            }
            else if (md.IsTableBlock)
            {
                var mdt = md as MarkdownParagraph.TableBlock;
                var header = mdt.headers.Option();
                var align = mdt.alignments;
                var rows = mdt.rows;
                foreach (var col in header) foreach (var s in Paragraphs2Words(col)) yield return s;
                foreach (var row in rows) foreach (var col in row) foreach (var s in Paragraphs2Words(col)) yield return s;
            }
            else if (md.IsEmbedParagraphs)
            {
            }
            else if (md.IsHorizontalRule)
            {
            }
            else if (md.IsInlineBlock)
            {
                var mdi = md as MarkdownParagraph.InlineBlock;
                foreach (var s in String2Words(mdi.code)) yield return s;
            }
            else if (md.IsLatexBlock)
            {
                var mdl = md as MarkdownParagraph.LatexBlock;
                foreach (var str in mdl.body) foreach (var s in String2Words(str)) yield return s;
            }
            else if (md.IsQuotedBlock)
            {
                var mdq = md as MarkdownParagraph.QuotedBlock;
                foreach (var s in Paragraphs2Words(mdq.paragraphs)) yield return s;
            }
            else if (md.IsSpan)
            {
                var mds = md as MarkdownParagraph.Span;
                foreach (var s in Spans2Words(mds.body)) yield return s;
            }
        }

        private static IEnumerable<string> String2Words(string src)
        {
            if (src == null) yield break;
            int i = 0;
            while (i < src.Length && !IsWordChar(src[i])) i++;
            while (i < src.Length)
            {
                string word = "";
                int spanStart = i;
                while (i < src.Length && IsWordChar(src[i])) { word += src[i]; i++; }
                yield return word;
                while (i < src.Length && !IsWordChar(src[i])) i++;
            }
        }

        public static Span FindSection(string src, SectionRef sr)
        {
            var hashes = new string('#', sr.Level);
            var sections = from section in FindSections(src)
                           where section.hashes == hashes
                           let i = LevenshteinCharDistance(section.title, sr.MarkdownTitle)
                           orderby i ascending
                           select section.span;
            return sections.FirstOrDefault();
        }

        private class SectionSpan
        {
            public string hashes, title;
            public Span span;
        }

        private static List<SectionSpan> FindSections(string src)
        {
            var sections = FindSectionsInner(src).ToList();
            for (int i = 0; i < sections.Count; i++)
            {
                var start = sections[i].span.Start;
                var nextStart = i == sections.Count - 1 ? src.Length : sections[i + 1].span.Start;
                sections[i].span = new Span(start, nextStart - start);
            }
            return sections;
        }

        private static IEnumerable<SectionSpan> FindSectionsInner(string src)
        {
            for (int i = 0; i < src.Length;)
            {
                // invariant: i is at the start of a line
                if (src[i] == '#')
                {
                    int lineSpanStart = i;
                    string line = "";
                    while (i < src.Length && src[i] != '\r' && src[i] != '\n') { line += src[i]; i++; }
                    while (i < src.Length && (src[i] == '\r' || src[i] == '\n')) i++;
                    var hashes = ""; while (line.StartsWith("#")) { hashes += line[0]; line = line.Substring(1); }
                    var title = line.TrimStart(' ');
                    yield return new SectionSpan { span = new Span(lineSpanStart, 0), hashes = hashes, title = title };
                }
                else
                {
                    while (i < src.Length && src[i] != '\r' && src[i] != '\n') i++;
                    while (i < src.Length && (src[i] == '\r' || src[i] == '\n')) i++;
                }
            }
        }

        private static int LevenshteinCharDistance(string x, string y) => LevenshteinDistance(x, y);

        private static int LevenshteinDistance<T>(IEnumerable<T> x, IEnumerable<T> y) where T : IEquatable<T>
        {
            if (x == null) throw new ArgumentNullException(nameof(x));
            if (y == null) throw new ArgumentNullException(nameof(y));
            IList<T> xx = x as IList<T> ?? new List<T>(x);
            IList<T> yy = y as IList<T> ?? new List<T>(y);
            if (xx.Count == 0) return yy.Count;
            if (yy.Count == 0) return xx.Count;

            // Rather than maintain an entire matrix (which would require O(n*m) space),
            // just store the current row and the next row, each of which has a length m+1,
            // so just O(m) space. Initialize the current row.
            int curRow = 0, nextRow = 1;
            int[][] rows = new int[][] { new int[yy.Count + 1], new int[yy.Count + 1] };
            for (int j = 0; j <= yy.Count; ++j) rows[curRow][j] = j;

            // For each virtual row (since we only have physical storage for two)
            for (int i = 1; i <= xx.Count; ++i)
            {
                // Fill in the values in the row
                rows[nextRow][0] = i;
                for (int j = 1; j <= yy.Count; ++j)
                {
                    int dist1 = rows[curRow][j] + 1;
                    int dist2 = rows[nextRow][j - 1] + 1;
                    int dist3 = rows[curRow][j - 1] + (xx[i - 1].Equals(yy[j - 1]) ? 0 : 1);
                    rows[nextRow][j] = Math.Min(dist1, Math.Min(dist2, dist3));
                }

                // Swap the current and next rows
                if (curRow == 0) { curRow = 1; nextRow = 0; }
                else { curRow = 0; nextRow = 1; }
            }

            return rows[curRow][yy.Count];
        }

        private static Span LevenshteinSearch<T>(IEnumerable<T> needle, IEnumerable<T> haystack) where T : IEquatable<T>
        {
            var s = LevenshteinSearchInner(needle, haystack);
            if (s == null) return null;
            var spanEnd = s.Length;
            s = LevenshteinSearchInner(needle.Reverse(), haystack.Reverse());
            var spanStart = haystack.Count() - s.Length;
            if (spanEnd < spanStart) return null;
            return new Span(spanStart, spanEnd - spanStart);
        }

        private static Span LevenshteinSearchInner<T>(IEnumerable<T> needle0, IEnumerable<T> haystack0) where T : IEquatable<T>
        {
            if (needle0 == null) throw new ArgumentNullException(nameof(needle0));
            if (haystack0 == null) throw new ArgumentNullException(nameof(haystack0));
            IList<T> needle = needle0 as IList<T> ?? new List<T>(needle0);
            IList<T> haystack = haystack0 as IList<T> ?? new List<T>(haystack0);
            if (needle.Count == 0) throw new ArgumentOutOfRangeException(nameof(needle0), "must be non-empty needle");
            if (haystack.Count == 0) throw new ArgumentOutOfRangeException(nameof(haystack0), "must be non-empty haystack");

            int curRow = 0, nextRow = 1;
            int[][] rows = new int[][] { new int[haystack.Count + 1], new int[haystack.Count + 1] };
            for (int j = 0; j <= haystack.Count; ++j) rows[curRow][j] = 0;

            for (int i = 1; i <= needle.Count; ++i)
            {
                rows[nextRow][0] = i;
                for (int j = 1; j <= haystack.Count; ++j)
                {
                    int dist1 = rows[curRow][j] + 1;
                    int dist2 = rows[nextRow][j - 1] + 1;
                    int dist3 = rows[curRow][j - 1] + (needle[i - 1].Equals(haystack[j - 1]) ? 0 : 1);
                    rows[nextRow][j] = Math.Min(dist1, Math.Min(dist2, dist3));
                }

                if (curRow == 0) { curRow = 1; nextRow = 0; }
                else { curRow = 0; nextRow = 1; }
            }

            var minScore = Enumerable.Min(rows[curRow]);
            var matches = rows[curRow].Select((score, i) => new { score, i }).Where(t => t.score == minScore).Select(t => t.i).ToArray();
            if (matches.Length != 1) return null;
            return new Span(0, matches[0]);
        }
    }
}