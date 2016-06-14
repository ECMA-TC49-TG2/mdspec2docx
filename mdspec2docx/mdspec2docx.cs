﻿using CSharp2Colorized;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FSharp.Markdown;
using Grammar2Html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;


class MarkdownSpec
{
    private string s;
    private IEnumerable<string> files;
    public Grammar Grammar = new Grammar();
    public List<SectionRef> Sections = new List<SectionRef>();
    public List<ProductionRef> Productions = new List<ProductionRef>();

    public class ProductionRef
    {
        public string Code;                  // the complete antlr code block in which it's found
        public List<string> ProductionNames; // all production names in it
        public string BookmarkName;          // _Grm00023
        public static int count = 1;

        public ProductionRef(string code, IEnumerable<Production> productions)
        {
            Code = code;
            ProductionNames = new List<string>(from p in productions where p.ProductionName != null select p.ProductionName);
            BookmarkName = $"_Grm{count:00000}"; count++;
        }
    }

    public class TermRef
    {
        public string Term;
        public string BookmarkName;
        public static int count = 1;

        public TermRef(string term)
        {
            Term = term;
            BookmarkName = $"_Trm{count:00000}"; count++;
        }
    }

    public class SectionRef
    {
        public string Number;        // "10.1.2"
        public string Title;         // "Goto Statement"
        public int Level;            // 1-based level, e.g. 3
        public string Url;           // statements.md#goto-statement
        public string BookmarkName;  // _Toc00023
        public static int count = 1;

        public SectionRef(MarkdownParagraph.Heading mdh, string filename)
        {
            Level = mdh.Item1;
            var spans = mdh.Item2;
            if (spans.Length == 1 && spans.First().IsLiteral) Title = mdunescape(spans.First() as MarkdownSpan.Literal).Trim();
            else if (spans.Length == 1 && spans.First().IsInlineCode) Title = (spans.First() as MarkdownSpan.InlineCode).Item.Trim();
            else throw new NotSupportedException("Heading must be a single literal/inlinecode");
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
        }
    }



    public static MarkdownSpec ReadString(string s)
    {
        var md = new MarkdownSpec { s = s };
        md.Init();
        return md;
    }

    public static MarkdownSpec ReadFiles(IEnumerable<string> files)
    {
        var md = new MarkdownSpec { files = files };
        md.Init();
        return md;
    }

    private void Init()
    {
        // (1) Add sections into the dictionary
        int h1 = 0, h2 = 0, h3 = 0, h4 = 0;
        string url = "", title = "";

        // (2) Turn all the antlr code blocks into a grammar
        var sbantlr = new StringBuilder();

        foreach (var src in Sources())
        {
            var filename = Path.GetFileName(src.Item1);
            var md = Markdown.Parse(src.Item2);

            foreach (var mdp in md.Paragraphs)
            {
                if (mdp.IsHeading)
                {
                    var sr = new SectionRef(mdp as MarkdownParagraph.Heading, filename);
                    if (sr.Level == 1) { h1 += 1; h2 = 0; h3 = 0; h4 = 0; sr.Number = $"{h1}"; }
                    if (sr.Level == 2) { h2 += 1; h3 = 0; h4 = 0; sr.Number = $"{h1}.{h2}"; }
                    if (sr.Level == 3) { h3 += 1; h4 = 0; sr.Number = $"{h1}.{h2}.{h3}"; }
                    if (sr.Level == 4) { h4 += 1; sr.Number = $"{h1}.{h2}.{h3}.{h4}"; }
                    if (sr.Level > 4) throw new NotSupportedException("Only support heading depths up to ####");
                    if (Sections.Any(s => s.Url == sr.Url)) throw new Exception($"Duplicate section title {sr.Url}");
                    Sections.Add(sr);
                    url = sr.Url;
                    title = sr.Title;
                }
                else if (mdp.IsCodeBlock)
                {
                    var mdc = mdp as MarkdownParagraph.CodeBlock;
                    string code = mdc.Item1, lang = mdc.Item2;
                    if (lang != "antlr") continue;
                    var g = Antlr.ReadString(code, "");
                    Productions.Add(new ProductionRef(code, g.Productions));
                    foreach (var p in g.Productions)
                    {
                        p.Link = url; p.LinkName = title;
                        if (p.ProductionName != null && Grammar.Productions.Any(dupe => dupe.ProductionName == p.ProductionName))
                        {
                            Console.WriteLine($"Duplicate grammar for {p.ProductionName}");
                        }
                        Grammar.Productions.Add(p);
                    }
                }
            }



        }
    }

    private IEnumerable<Tuple<string, string>> Sources()
    {
        if (s != null) yield return Tuple.Create("", BugWorkaroundEncode(s));
        foreach (var fn in files ?? new string[] { })
        {
            yield return Tuple.Create(fn, BugWorkaroundEncode(File.ReadAllText(fn)));
        }
    }

    private static string BugWorkaroundEncode(string src)
    {
        var lines = src.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

        // https://github.com/tpetricek/FSharp.formatting/issues/388
        // The markdown parser doesn't recognize | inside inlinecode inside table
        // To work around that, we'll encode this |, then decode it later
        for (int li = 0; li < lines.Length; li++)
        {
            if (!lines[li].StartsWith("|")) continue;
            var codes = lines[li].Split('`');
            for (int ci = 1; ci < codes.Length; ci += 2)
            {
                codes[ci] = codes[ci].Replace("|", "ceci_n'est_pas_une_pipe");
            }
            lines[li] = string.Join("`", codes);
        }

        // https://github.com/tpetricek/FSharp.formatting/issues/347
        // The markdown parser overly indents a level1 paragraph if it follows a level2 bullet
        // To work around that, we'll call out now, then unindent it later
        var state = 0; // 1=found.level1, 2=found.level2
        for (int li = 0; li < lines.Length - 1; li++)
        {
            if (lines[li].StartsWith("*  "))
            {
                state = 1;
                if (string.IsNullOrWhiteSpace(lines[li + 1])) li++;
            }
            else if ((state == 1 || state == 2) && lines[li].StartsWith("   * "))
            {
                state = 2;
                if (string.IsNullOrWhiteSpace(lines[li + 1])) li++;
            }
            else if (state == 2 && lines[li].StartsWith("      ") && lines[li].Length > 6 && lines[li][6] != ' ')
            {
                state = 2;
                if (string.IsNullOrWhiteSpace(lines[li + 1])) li++;
            }
            else if (state == 2 && lines[li].StartsWith("   ") && lines[li].Length > 3 && lines[li][3] != ' ')
            {
                lines[li] = "   ceci-n'est-pas-une-indent" + lines[li].Substring(3);
                state = 0;
            }
            else
            {
                state = 0;
            }
        }

        src = string.Join("\r\n", lines);

        // https://github.com/tpetricek/FSharp.formatting/issues/390
        // The markdown parser doesn't recognize bullet-chars inside codeblocks inside lists
        // To work around that, we'll prepend the line with stuff, and remove it later
        var codeblocks = src.Split(new[] { "\r\n    ```" }, StringSplitOptions.None);
        for (int cbi = 1; cbi < codeblocks.Length; cbi += 2)
        {
            var s = codeblocks[cbi];
            s = s.Replace("\r\n    *", "\r\n    ceci_n'est_pas_une_*");
            s = s.Replace("\r\n    +", "\r\n    ceci_n'est_pas_une_+");
            s = s.Replace("\r\n    -", "\r\n    ceci_n'est_pas_une_-");
            codeblocks[cbi] = s;
        }

        return string.Join("\r\n    ```", codeblocks);
    }

    private static string BugWorkaroundDecode(string s)
    {
        // This function should be alled on all inline-code and code blocks
        s = s.Replace("ceci_n'est_pas_une_pipe", "|");
        s = s.Replace("ceci_n'est_pas_une_", "");
        return s;
    }

    private static int BugWorkaroundIndent(ref MarkdownParagraph mdp, int level)
    {
        if (!mdp.IsParagraph) return level;
        var p = mdp as MarkdownParagraph.Paragraph;
        var spans = p.Item;
        if (spans.Count() == 0 || !spans[0].IsLiteral) return level;
        var literal = spans[0] as MarkdownSpan.Literal;
        if (!literal.Item.StartsWith("ceci-n'est-pas-une-indent")) return level;
        //
        var literal2 = MarkdownSpan.NewLiteral(literal.Item.Substring(25));
        var spans2 = Microsoft.FSharp.Collections.FSharpList<MarkdownSpan>.Cons(literal2, spans.Tail);
        var p2 = MarkdownParagraph.NewParagraph(spans2);
        mdp = p2;
        return 0;
    }

    bool FindToc(Body body, out int ifirst, out int iLast, out string instr, out Paragraph secBreak)
    {
        ifirst = -1; iLast = -1; instr = null; secBreak = null;

        for (int i = 0; i < body.ChildElements.Count; i++)
        {
            var p = body.ChildElements.GetItem(i) as Paragraph;
            if (p == null) continue;

            // The TOC might be a simple field
            var sf = p.OfType<SimpleField>().FirstOrDefault();
            if (sf != null && sf.Instruction.Value.Contains("TOC"))
            {
                if (ifirst != -1) throw new Exception("Found start of TOC and then another simple TOC");
                ifirst = i; iLast = i; instr = sf.Instruction.Value;
                break;
            }

            // or it might be a complex field
            var runElements = (from r in p.OfType<Run>() from e in r select e).ToList();
            var f1 = runElements.FindIndex(f => f is FieldChar && (f as FieldChar).FieldCharType.Value == FieldCharValues.Begin);
            var f2 = runElements.FindIndex(f => f is FieldCode && (f as FieldCode).Text.Contains("TOC"));
            var f3 = runElements.FindIndex(f => f is FieldChar && (f as FieldChar).FieldCharType.Value == FieldCharValues.Separate);
            var f4 = runElements.FindIndex(f => f is FieldChar && (f as FieldChar).FieldCharType.Value == FieldCharValues.End);

            if (f1 != -1 && f2 != -1 && f3 != -1 && f2 > f1 && f3 > f2)
            {
                if (ifirst != -1) throw new Exception("Found start of TOC and then another start of TOC");
                ifirst = i; instr = (runElements[f2] as FieldCode).Text;
            }
            if (f4 != -1 && f4 > f1 && f4 > f2 && f4 > f3)
            {
                iLast = i;
                if (ifirst != -1) break;
            }
        }

        if (ifirst == -1) return false;
        if (iLast == -1) throw new Exception("Found start of TOC field, but not end");
        for (int i = ifirst; i <= iLast; i++)
        {
            var p = body.ChildElements.GetItem(i) as Paragraph;
            if (p == null) continue;
            var sp = p.ParagraphProperties.OfType<SectionProperties>().FirstOrDefault();
            if (sp == null) continue;
            if (i != iLast) throw new Exception("Found section break within TOC field");
            secBreak = new Paragraph(new Run(new Text(""))) { ParagraphProperties = new ParagraphProperties(sp.CloneNode(true)) };
        }
        return true;
    }


    public class StringLengthComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            if (x.Length > y.Length) return -1;
            if (x.Length < y.Length) return 1;
            return string.Compare(x, y);
        }
    }

    public void WriteFile(string basedOn, string fn)
    {
        using (var templateDoc = WordprocessingDocument.Open(basedOn, false))
        using (var resultDoc = WordprocessingDocument.Create(fn, WordprocessingDocumentType.Document))
        {

            foreach (var part in templateDoc.Parts) resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
            var body = resultDoc.MainDocumentPart.Document.Body;

            // We have to find the TOC, if one exists, and replace it...
            var tocFirst = -1;
            var tocLast = -1;
            var tocInstr = "";
            var tocSec = null as Paragraph;

            if (FindToc(body, out tocFirst, out tocLast, out tocInstr, out tocSec))
            {
                var tocRunFirst = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin },
                                          new FieldCode { Text = tocInstr, Space = SpaceProcessingModeValues.Preserve },
                                          new FieldChar { FieldCharType = FieldCharValues.Separate });
                var tocRunLast = new Run(new FieldChar { FieldCharType = FieldCharValues.End });
                //
                for (int i = tocLast; i >= tocFirst; i--) body.RemoveChild(body.ChildElements[i]);
                var afterToc = body.ChildElements[tocFirst];
                //
                for (int i = 0; i < Sections.Count; i++)
                {
                    var section = Sections[i];
                    if (section.Level > 2) continue;
                    var p = new Paragraph();
                    if (i == 0) p.AppendChild(tocRunFirst);
                    p.AppendChild(new Hyperlink(new Run(new Text(section.Number + " " + section.Title))) { Anchor = section.BookmarkName });
                    if (i == Sections.Count - 1) p.AppendChild(tocRunLast);
                    p.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = $"TOC{section.Level}" });
                    body.InsertBefore(p, afterToc);
                }
                if (tocSec != null) body.InsertBefore(tocSec, afterToc);
            }

            var maxBookmarkId = new StrongBox<int>(1 + body.Descendants<BookmarkStart>().Max(bookmark => int.Parse(bookmark.Id)));
            var terms = new SortedList<string, TermRef>();
            var italics = new List<Tuple<string, string>>();
            foreach (var src in Sources())
            {
                var converter = new MarkdownConverter
                {
                    Mddoc = Markdown.Parse(src.Item2),
                    Filename = Path.GetFileName(src.Item1),
                    Wdoc = resultDoc,
                    Sections = Sections.ToDictionary(sr => sr.Url),
                    Productions = Productions,
                    Terms = terms,
                    Italics = italics,
                    MaxBookmarkId = maxBookmarkId
                };
                foreach (var p in converter.Paragraphs())
                {
                    body.AppendChild(p);
                }
            }

            // I wonder if there were any oddities? ...
            // Terms that were referenced before their definition?
            var termset = new HashSet<string>(terms.Keys);
            var italicset = new HashSet<string>(italics.Where(i => i.Item2 == "italic").Select(i => i.Item1));
            italicset.IntersectWith(termset);
            if (italicset.Any())
            {
                Console.WriteLine("ERROR - these terms were used before their definition: " + string.Join(",", italicset));
            }

            // Terms that are also production names?
            var productionset = new HashSet<string>(Grammar.Productions.Where(p => p.ProductionName != null).Select(p => p.ProductionName));
            productionset.IntersectWith(termset);
            if (productionset.Any())
            {
                Console.WriteLine("ERROR - these terms are also production names: " + string.Join(",", productionset));
            }

            // Terms that were defined but never used?
            var termrefset = new HashSet<string>(italics.Where(i => i.Item2 == "term").Select(i => i.Item1));
            termset.RemoveWhere(t => termrefset.Contains(t));
            if (termset.Any())
            {
                Console.WriteLine("ERROR - these terms are defined but never used: " + string.Join(",", termset));
            }

            // Which single-word production-names appear in italics?
            var italicproductionset = new HashSet<string>(italics.Where(i => !i.Item1.Contains("_") && i.Item2 == "production").Select(i => i.Item1));
            var italicproductions = string.Join(",", italicproductionset);

            // What are the single-word production names that don't appear in italics?
            var otherproductionset = new HashSet<string>(Grammar.Productions.Where(p => p.ProductionName != null && !p.ProductionName.Contains("_") && !italicproductionset.Contains(p.ProductionName)).Select(p => p.ProductionName));
            var otherproductions = string.Join(",", otherproductionset);

        }
    }

    static string mdunescape(MarkdownSpan.Literal literal) =>
        literal.Item.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">").Replace("&reg;", "®");


    private class MarkdownConverter
    {
        public MarkdownDocument Mddoc;
        public WordprocessingDocument Wdoc;
        public Dictionary<string, SectionRef> Sections;
        public List<ProductionRef> Productions;
        public SortedList<string,TermRef> Terms;
        public List<Tuple<string, string>> Italics;
        public StrongBox<int> MaxBookmarkId;
        public string Filename;
        public string CurrentSection;

        public IEnumerable<OpenXmlCompositeElement> Paragraphs()
            => Paragraphs2Paragraphs(Mddoc.Paragraphs);

        IEnumerable<OpenXmlCompositeElement> Paragraphs2Paragraphs(IEnumerable<MarkdownParagraph> pars)
        {
            foreach (var md in pars) foreach (var p in Paragraph2Paragraphs(md)) yield return p;
        }


        IEnumerable<OpenXmlCompositeElement> Paragraph2Paragraphs(MarkdownParagraph md)
        {
            if (md.IsHeading)
            {
                var mdh = md as MarkdownParagraph.Heading;
                var level = mdh.Item1;
                var spans = mdh.Item2;
                var sr = Sections[new SectionRef(mdh, Filename).Url];
                var props = new ParagraphProperties(new ParagraphStyleId() { Val = $"Heading{level}" });
                var p = new Paragraph { ParagraphProperties = props };
                MaxBookmarkId.Value += 1;
                p.AppendChild(new BookmarkStart { Name = sr.BookmarkName, Id = MaxBookmarkId.Value.ToString() });
                p.Append(Spans2Elements(spans));
                p.AppendChild(new BookmarkEnd { Id = MaxBookmarkId.Value.ToString() });
                yield return p;
                //
                var i = sr.Url.IndexOf("#");
                CurrentSection = $"{sr.Url.Substring(0, i)} {new string('#', level)} {sr.Title} [{sr.Number}]";
                //Console.WriteLine(CurrentSection); // new string(' ', level * 4 - 4) + sr.Number + " " + sr.Title);
                yield break;
            }

            else if (md.IsParagraph)
            {
                var mdp = md as MarkdownParagraph.Paragraph;
                var spans = mdp.Item;
                yield return new Paragraph(Spans2Elements(spans));
                yield break;
            }

            else if (md.IsListBlock)
            {
                var mdl = md as MarkdownParagraph.ListBlock;
                var flat = FlattenList(mdl);

                // Let's figure out what kind of list it is - ordered or unordered? nested?
                var format0 = new[] { "1", "1", "1", "1" };
                foreach (var item in flat) format0[item.Level] = (item.IsBulletOrdered ? "1" : "o");
                var format = string.Join("", format0);

                var numberingPart = Wdoc.MainDocumentPart.NumberingDefinitionsPart ?? Wdoc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");
                if (numberingPart.Numbering == null) numberingPart.Numbering = new Numbering();

                Func<int, bool, Level> createLevel;
                createLevel = (level, isOrdered) =>
                {
                    var numformat = NumberFormatValues.Bullet;
                    var levelText = new[] { "·", "o", "·", "o" }[level];
                    if (isOrdered && level == 0) { numformat = NumberFormatValues.Decimal; levelText = "%1."; }
                    if (isOrdered && level == 1) { numformat = NumberFormatValues.LowerLetter; levelText = "%2."; }
                    if (isOrdered && level == 2) { numformat = NumberFormatValues.LowerRoman; levelText = "%3."; }
                    if (isOrdered && level == 3) { numformat = NumberFormatValues.LowerRoman; levelText = "%4."; }
                    var r = new Level { LevelIndex = level };
                    r.Append(new StartNumberingValue { Val = 1 });
                    r.Append(new NumberingFormat { Val = numformat });
                    r.Append(new LevelText { Val = levelText });
                    r.Append(new ParagraphProperties(new Indentation { Left = (540 + 360 * level).ToString(), Hanging = "360" }));
                    if (levelText == "·") r.Append(new NumberingSymbolRunProperties(new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol", EastAsia = "Times new Roman", ComplexScript = "Times new Roman" }));
                    if (levelText == "o") r.Append(new NumberingSymbolRunProperties(new RunFonts { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" }));
                    return r;
                };
                var level0 = createLevel(0, format[0] == '1');
                var level1 = createLevel(1, format[1] == '1');
                var level2 = createLevel(2, format[2] == '1');
                var level3 = createLevel(3, format[3] == '1');

                var abstracts = numberingPart.Numbering.OfType<AbstractNum>().Select(an => an.AbstractNumberId.Value).ToList();
                var aid = (abstracts.Count == 0 ? 1 : abstracts.Max() + 1);
                var aabstract = new AbstractNum(new MultiLevelType() {Val = MultiLevelValues.Multilevel}, level0, level1, level2, level3) {AbstractNumberId = aid};
                numberingPart.Numbering.InsertAt(aabstract, 0);

                var instances = numberingPart.Numbering.OfType<NumberingInstance>().Select(ni => ni.NumberID.Value);
                var nid = (instances.Count() == 0 ? 1 : instances.Max() + 1);
                var numInstance = new NumberingInstance(new AbstractNumId { Val = aid }) { NumberID = nid };
                numberingPart.Numbering.AppendChild(numInstance);

                // We'll also figure out the indentation(for the benefit of those paragraphs that should be
                // indendent with the list but aren't numbered). I'm not sure what the indent comes from.
                // in the docx, each AbstractNum that I created has an indent for each of its levels,
                // defaulted at 900, 1260, 1620, ... but I can't see where in the above code that's created?
                Func<int,string> calcIndent = level => (540 + level * 360).ToString();

                foreach (var item in flat)
                {
                    var content = item.Paragraph;
                    if (content.IsParagraph || content.IsSpan)
                    {
                        var spans = (content.IsParagraph ? (content as MarkdownParagraph.Paragraph).Item : (content as MarkdownParagraph.Span).Item);
                        if (item.HasBullet) yield return new Paragraph(Spans2Elements(spans)) { ParagraphProperties = new ParagraphProperties(new NumberingProperties(new ParagraphStyleId { Val = "ListParagraph" }, new NumberingLevelReference { Val = item.Level }, new NumberingId { Val = nid })) };
                        else yield return new Paragraph(Spans2Elements(spans)) { ParagraphProperties = new ParagraphProperties(new Indentation { Left = calcIndent(item.Level) }) };
                    }
                    else if (content.IsQuotedBlock || content.IsCodeBlock)
                    {
                        foreach (var p in Paragraph2Paragraphs(content))
                        {
                            var props = p.GetFirstChild<ParagraphProperties>();
                            if (props == null) { props = new ParagraphProperties(); p.InsertAt(props, 0); }
                            var indent = props?.GetFirstChild<Indentation>();
                            if (indent == null) { indent = new Indentation(); props.Append(indent); }
                            indent.Left = calcIndent(item.Level);
                            yield return p;
                        }
                    }
                    else if (content.IsTableBlock)
                    {
                        foreach (var p in Paragraph2Paragraphs(content))
                        {
                            var table = p as Table;
                            if (table == null) { yield return p; continue; }
                            var tprops = table.GetFirstChild<TableProperties>();
                            var tindent = tprops?.GetFirstChild<TableIndentation>();
                            if (tindent == null) throw new Exception("Ooops! Table is missing indentation");
                            tindent.Width = int.Parse(calcIndent(item.Level));
                            yield return table;
                        }
                    }
                    else
                    {
                        throw new Exception("Unexpected item in list");
                    }
                }
            }

            else if (md.IsCodeBlock)
            {
                var mdc = md as MarkdownParagraph.CodeBlock;
                var code = mdc.Item1;
                var lang = mdc.Item2;
                code = BugWorkaroundDecode(code);
                var runs = new List<Run>();
                var onFirstLine = true;
                IEnumerable<ColorizedLine> lines;
                if (lang == "csharp" || lang == "c#" || lang == "cs") lines = Colorize.CSharp(code);
                else if (lang == "vb" || lang == "vbnet" || lang == "vb.net") lines = Colorize.VB(code);
                else if (lang == "" || lang == "xml") lines = Colorize.PlainText(code);
                else if (lang == "antlr") lines = Antlr.ColorizeAntlr(code);
                else throw new NotSupportedException($"unrecognized language {lang}");
                foreach (var line in lines)
                {
                    if (onFirstLine) onFirstLine = false; else runs.Add(new Run(new Break()));
                    foreach (var word in line.Words)
                    {
                        var run = new Run();
                        var props = new RunProperties();
                        if (word.Red != 0 || word.Green != 0 || word.Blue != 0) props.Append(new Color { Val = $"{word.Red:X2}{word.Green:X2}{word.Blue:X2}" });
                        if (word.IsItalic) props.Append(new Italic());
                        if (props.HasChildren) run.Append(props);
                        run.Append(new Text(word.Text) { Space = SpaceProcessingModeValues.Preserve });
                        runs.Add(run);
                    }
                }
                if (lang == "antlr")
                {
                    var p = new Paragraph() { ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Grammar" }) };
                    var prodref = Productions.Single(prod => prod.Code == code);
                    MaxBookmarkId.Value += 1;
                    p.AppendChild(new BookmarkStart { Name = prodref.BookmarkName, Id = MaxBookmarkId.Value.ToString() });
                    p.Append(runs);
                    p.AppendChild(new BookmarkEnd { Id = MaxBookmarkId.Value.ToString() });
                    yield return p;
                }
                else
                {
                    var p = new Paragraph() { ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Code" }) };
                    p.Append(runs);
                    yield return p;
                }
            }

            else if (md.IsQuotedBlock)
            {
                throw new NotSupportedException("Quoted blocks not supported");
            }

            else if (md.IsTableBlock)
            {
                var mdt = md as MarkdownParagraph.TableBlock;
                var header = mdt.Item1.Option();
                var align = mdt.Item2;
                var rows = mdt.Item3;
                var table = new Table();
                if (header == null) Console.WriteLine("ERROR - github requires all tables to have header rows");
                if (!header.Any(cell => cell.Length > 0)) header = null; // even if Github requires an empty header, we can at least cull it from Docx
                var tstyle = new TableStyle { Val = "TableGrid" };
                var tindent = new TableIndentation { Width = 360, Type = TableWidthUnitValues.Dxa };
                var tborders = new TableBorders();
                tborders.TopBorder = new TopBorder { Val = BorderValues.Single };
                tborders.BottomBorder = new BottomBorder { Val = BorderValues.Single };
                tborders.LeftBorder = new LeftBorder { Val = BorderValues.Single };
                tborders.RightBorder = new RightBorder { Val = BorderValues.Single };
                tborders.InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single };
                tborders.InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Single };
                var tcellmar = new TableCellMarginDefault();
                tcellmar.Append();
                table.Append(new TableProperties(tstyle, tindent, tborders));
                var ncols = align.Length;
                for (int irow = -1; irow < rows.Length; irow++)
                {
                    if (irow == -1 && header == null) continue;
                    var mdrow = (irow == -1 ? header : rows[irow]);
                    var row = new TableRow();
                    for (int icol = 0; icol < Math.Min(ncols, mdrow.Length); icol++)
                    {
                        var mdcell = mdrow[icol];
                        var cell = new TableCell();
                        var pars = Paragraphs2Paragraphs(mdcell).ToList();
                        for (int ip = 0; ip < pars.Count; ip++)
                        {
                            var p = pars[ip] as Paragraph;
                            if (p == null) { cell.Append(pars[ip]); continue;}
                            var props = new ParagraphProperties(new ParagraphStyleId { Val = "TableCellNormal" });
                            if (align[icol].IsAlignCenter) props.Append(new Justification { Val = JustificationValues.Center });
                            if (align[icol].IsAlignRight) props.Append(new Justification { Val = JustificationValues.Right });
                            p.InsertAt(props, 0);
                            cell.Append(pars[ip]);
                        }
                        if (pars.Count == 0) cell.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "0" }), new Run(new Text(""))));
                        row.Append(cell);
                    }
                    table.Append(row);
                }
                yield return new Paragraph(new Run(new Text(""))) { ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "TableLineBefore" }) };
                yield return table;
                yield return new Paragraph(new Run(new Text(""))) { ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "TableLineAfter" }) };
            }
            else
            {
                yield return new Paragraph(new Run(new Text($"[{md.GetType().Name}]")));
            }
        }

        class FlatItem
        {
            public int Level;
            public bool HasBullet;
            public bool IsBulletOrdered;
            public MarkdownParagraph Paragraph;
        }

        IEnumerable<FlatItem> FlattenList(MarkdownParagraph.ListBlock md)
        {
            var flat = FlattenList(md, 0).ToList();
            var isOrdered = new Dictionary<int, bool>();
            foreach (var item in flat)
            {
                var level = item.Level;
                var isItemOrdered = item.IsBulletOrdered;
                var content = item.Paragraph;
                if (isOrdered.ContainsKey(level) && isOrdered[level] != isItemOrdered) throw new NotImplementedException("List can't mix ordered and unordered items at same level");
                isOrdered[level] = isItemOrdered;
                if (level > 3) throw new Exception("Can't have more than 4 levels in a list");
            }
            return flat;
        }

        IEnumerable<FlatItem> FlattenList(MarkdownParagraph.ListBlock md, int level)
        {
            var isOrdered = md.Item1.IsOrdered;
            var items = md.Item2;
            foreach (var mdpars in items)
            {
                var isFirstParagraph = true;
                foreach (var mdp in mdpars)
                {
                    var wasFirstParagraph = isFirstParagraph; isFirstParagraph = false;

                    if (mdp.IsParagraph || mdp.IsSpan)
                    {
                        var mdp1 = mdp;
                        var buglevel = BugWorkaroundIndent(ref mdp1, level);
                        yield return new FlatItem { Level = buglevel, HasBullet = wasFirstParagraph, IsBulletOrdered = isOrdered, Paragraph = mdp1 };
                    }
                    else if (mdp.IsQuotedBlock || mdp.IsCodeBlock)
                    {
                        yield return new FlatItem { Level = level, HasBullet = false, IsBulletOrdered = isOrdered, Paragraph = mdp };
                    }
                    else if (mdp.IsListBlock)
                    {
                        foreach (var subitem in FlattenList(mdp as MarkdownParagraph.ListBlock, level + 1)) yield return subitem;
                    }
                    else if (mdp.IsTableBlock)
                    {
                        yield return new FlatItem { Level = level, HasBullet = false, IsBulletOrdered = isOrdered, Paragraph = mdp };
                    }
                    else
                    {
                        throw new NotImplementedException("nothing fancy allowed in lists");
                    }
                }
            }
        }


        IEnumerable<OpenXmlElement> Spans2Elements(IEnumerable<MarkdownSpan> mds, bool nestedSpan = false)
        {
            foreach (var md in mds) foreach (var e in Span2Elements(md, nestedSpan)) yield return e;
        }

        IEnumerable<OpenXmlElement> Span2Elements(MarkdownSpan md, bool nestedSpan = false)
        {
            if (md.IsLiteral)
            {
                var mdl = md as MarkdownSpan.Literal;
                var s = mdunescape(mdl);
                foreach (var r in Literal2Elements(s, nestedSpan)) yield return r;
            }

            else if (md.IsStrong || md.IsEmphasis)
            {
                IEnumerable<MarkdownSpan> spans = (md.IsStrong ? (md as MarkdownSpan.Strong).Item : (md as MarkdownSpan.Emphasis).Item);

                // Workaround for https://github.com/tpetricek/FSharp.formatting/issues/389 - the markdown parser
                // turns *this_is_it* into a nested Emphasis["this", Emphasis["is"], "it"] instead of Emphasis["this_is_it"]
                // What we'll do is preprocess it into Emphasis["this_is_it"]
                if (md.IsEmphasis)
                {
                    var spans2 = spans.Select(s =>
                    {
                        var _ = "";
                        if (s.IsEmphasis) { s = (s as MarkdownSpan.Emphasis).Item.Single(); _ = "_"; }
                        if (s.IsLiteral) return _ + (s as MarkdownSpan.Literal).Item + _;
                        throw new NotSupportedException("something odd inside emphasis");
                    });
                    spans = new List<MarkdownSpan>() { MarkdownSpan.NewLiteral(string.Join("", spans2)) };
                }

                // Convention is that ***term*** is used to define a term.
                // That's parsed as Strong, which contains Emphasis, which contains one Literal
                string literal = null;
                TermRef termdef = null;
                if (!nestedSpan && md.IsStrong && spans.Count() == 1 && spans.First().IsEmphasis)
                {
                    var spans2 = (spans.First() as MarkdownSpan.Emphasis).Item;
                    if (spans2.Count() == 1 && spans2.First().IsLiteral)
                    {
                        literal = (spans2.First() as MarkdownSpan.Literal).Item;
                        termdef = new TermRef(literal);
                        if (Terms.ContainsKey(literal)) Console.WriteLine($"ERROR - term multiple definition - '{literal}'");
                        else Terms.Add(literal,termdef);
                    }
                }

                // Convention inside our specs is that emphasis only ever contains literals,
                // either to emphasis some human-text or to refer to an ANTLR-production
                ProductionRef prodref = null;
                if (!nestedSpan && md.IsEmphasis && (spans.Count() != 1 || !spans.First().IsLiteral)) throw new NotSupportedException("something odd inside emphasis");
                if (!nestedSpan && md.IsEmphasis && spans.Count() == 1 && spans.First().IsLiteral)
                {
                    literal = (spans.First() as MarkdownSpan.Literal).Item;
                    prodref = Productions.FirstOrDefault(pr => pr.ProductionNames.Contains(literal));
                    Italics.Add(Tuple.Create(literal, prodref != null ? "production" : "italic"));
                }

                if (prodref != null)
                {
                    var props = new RunProperties(new Color { Val = "6A5ACD" }, new Underline { Val=UnderlineValues.Single });
                    var run = new Run(new Text(literal) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props };
                    var link = new Hyperlink(run) { Anchor = prodref.BookmarkName };
                    yield return link;
                }
                else if (termdef != null)
                {
                    MaxBookmarkId.Value += 1;
                    yield return new BookmarkStart { Name = termdef.BookmarkName, Id = MaxBookmarkId.Value.ToString() };
                    var props = new RunProperties(new Italic(), new Bold());
                    yield return new Run(new Text(literal) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props };
                    yield return new BookmarkEnd { Id = MaxBookmarkId.Value.ToString() };
                }
                else
                {
                    foreach (var e in Spans2Elements(spans, true))
                    {
                        var style = (md.IsStrong ? new Bold() as OpenXmlElement : new Italic());
                        var run = e as Run;
                        if (run != null) run.InsertAt(new RunProperties(style), 0);
                        yield return e;
                    }
                }
            }

            else if (md.IsInlineCode)
            {
                var mdi = md as MarkdownSpan.InlineCode;
                var code = mdi.Item;

                var txt = new Text(BugWorkaroundDecode(code)) { Space = SpaceProcessingModeValues.Preserve };
                var props = new RunProperties(new RunStyle { Val = "CodeEmbedded" });
                var run = new Run(txt) { RunProperties = props };
                yield return run;
            }

            else if (md.IsDirectLink || md.IsIndirectLink)
            {
                IEnumerable<MarkdownSpan> spans;
                string url = "", alt = "";
                if (md.IsDirectLink)
                {
                    var mddl = md as MarkdownSpan.DirectLink;
                    spans = mddl.Item1;
                    url = mddl.Item2.Item1;
                    alt = mddl.Item2.Item2.Option();
                }
                else
                {
                    var mdil = md as MarkdownSpan.IndirectLink;
                    var original = mdil.Item2;
                    var id = mdil.Item3;
                    spans = mdil.Item1;
                    if (Mddoc.DefinedLinks.ContainsKey(id))
                    {
                        url = Mddoc.DefinedLinks[id].Item1;
                        alt = Mddoc.DefinedLinks[id].Item2.Option();
                    }
                }

                var anchor = "";
                if (spans.Count() == 1 && spans.First().IsLiteral) anchor = mdunescape(spans.First() as MarkdownSpan.Literal);
                else if (spans.Count() == 1 && spans.First().IsInlineCode) anchor = (spans.First() as MarkdownSpan.InlineCode).Item;
                else throw new NotImplementedException("Link anchor must be Literal or InlineCode, not " + md.ToString());

                if (Sections.ContainsKey(url))
                {
                    var section = Sections[url];
                    if (anchor != section.Title) throw new Exception($"Mismatch: link anchor is '{anchor}', should be '{section.Title}'");
                    var txt = new Text("§" + section.Number) { Space = SpaceProcessingModeValues.Preserve };
                    var run = new Hyperlink(new Run(txt)) { Anchor = section.BookmarkName };
                    yield return run;
                }
                else if (url.StartsWith("http:") || url.StartsWith("https:"))
                {
                    var style = new RunStyle { Val = "Hyperlink" };
                    var hyperlink = new Hyperlink { DocLocation = url, Tooltip = alt };
                    foreach (var element in Spans2Elements(spans))
                    {
                        var run = element as Run;
                        if (run != null) run.InsertAt(new RunProperties(style), 0);
                        hyperlink.AppendChild(run);
                    }
                    yield return hyperlink;
                }
                else
                {
                    throw new Exception("Absent hyperlink in " + md.ToString());
                }
            }

            else if (md.IsHardLineBreak)
            {
                // I've only ever seen this arise from dodgy markdown parsing, so I'll ignore it...
            }

            else
            {
                yield return new Run(new Text($"[{md.GetType().Name}]"));
            }
        }


        IEnumerable<OpenXmlElement> Literal2Elements(string literal, bool isNested)
        {
            foreach (var kv in Terms)
            {
                if (isNested) break;
                var i = literal.IndexOf(kv.Key);
                if (i == -1) continue;
                var left = literal.Substring(0, i);
                var right = literal.Substring(i + kv.Key.Length);
                var termref = kv.Value;
                Italics.Add(Tuple.Create(kv.Key, "term"));

                foreach (var r in Literal2Elements(left, isNested)) yield return r;
                //
                var props = new RunProperties(new Underline { Val = UnderlineValues.Dotted, Color = "4BACC6" });
                var run = new Run(new Text(kv.Key) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props };
                var link = new Hyperlink(run) { Anchor = termref.BookmarkName };
                yield return link;
                //
                foreach (var r in Literal2Elements(right, isNested)) yield return r;
                yield break;
            }

            yield return new Run(new Text(literal) { Space = SpaceProcessingModeValues.Preserve });

        }


    }

}


