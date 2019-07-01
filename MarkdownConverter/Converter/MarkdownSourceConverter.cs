using CSharp2Colorized;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FSharp.Formatting.Common;
using FSharp.Markdown;
using MarkdownConverter.Grammar;
using MarkdownConverter.Spec;
using Microsoft.FSharp.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace MarkdownConverter.Converter
{
    internal class MarkdownSourceConverter
    {
        public MarkdownDocument Mddoc;
        public WordprocessingDocument Wdoc;
        public Dictionary<string, SectionRef> Sections;
        public List<ProductionRef> Productions;
        public Dictionary<string, TermRef> Terms;
        public List<string> TermKeys;
        public List<ItalicUse> Italics;
        public StrongBox<int> MaxBookmarkId;
        public string Filename;
        public string CurrentSection;
        public Reporter Report;
        public ColorizerCache Colorizer;

        public IEnumerable<OpenXmlCompositeElement> Paragraphs()
            => Paragraphs2Paragraphs(Mddoc.Paragraphs);

        IEnumerable<OpenXmlCompositeElement> Paragraphs2Paragraphs(IEnumerable<MarkdownParagraph> pars)
        {
            foreach (var md in pars) foreach (var p in Paragraph2Paragraphs(md)) yield return p;
        }


        IEnumerable<OpenXmlCompositeElement> Paragraph2Paragraphs(MarkdownParagraph md)
        {
            Report.CurrentParagraph = md;
            if (md.IsHeading)
            {
                var mdh = md as MarkdownParagraph.Heading;
                var level = mdh.size;
                var spans = mdh.body;
                var sr = Sections[new SectionRef(mdh, Filename).Url];
                Report.CurrentSection = sr;
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
                Report.Log(CurrentSection);
                yield break;
            }

            else if (md.IsParagraph)
            {
                var mdp = md as MarkdownParagraph.Paragraph;
                var spans = mdp.body;
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
                var aabstract = new AbstractNum(new MultiLevelType() { Val = MultiLevelValues.Multilevel }, level0, level1, level2, level3) { AbstractNumberId = aid };
                numberingPart.Numbering.InsertAt(aabstract, 0);

                var instances = numberingPart.Numbering.OfType<NumberingInstance>().Select(ni => ni.NumberID.Value);
                var nid = (instances.Count() == 0 ? 1 : instances.Max() + 1);
                var numInstance = new NumberingInstance(new AbstractNumId { Val = aid }) { NumberID = nid };
                numberingPart.Numbering.AppendChild(numInstance);

                // We'll also figure out the indentation(for the benefit of those paragraphs that should be
                // indendent with the list but aren't numbered). I'm not sure what the indent comes from.
                // in the docx, each AbstractNum that I created has an indent for each of its levels,
                // defaulted at 900, 1260, 1620, ... but I can't see where in the above code that's created?
                Func<int, string> calcIndent = level => (540 + level * 360).ToString();

                foreach (var item in flat)
                {
                    var content = item.Paragraph;
                    if (content.IsParagraph || content.IsSpan)
                    {
                        var spans = (content.IsParagraph ? (content as MarkdownParagraph.Paragraph).body : (content as MarkdownParagraph.Span).body);
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
                        Report.Error("MD08", $"Unexpected item in list '{content.GetType().Name}'");
                    }
                }
            }

            else if (md.IsCodeBlock)
            {
                var mdc = md as MarkdownParagraph.CodeBlock;
                var code = mdc.code;
                var lang = mdc.language;
                code = BugWorkaroundDecode(code);
                var runs = new List<Run>();
                var onFirstLine = true;
                IEnumerable<ColorizedLine> lines;
                if (lang == "csharp" || lang == "c#" || lang == "cs") lines = Colorizer.Colorize("cs", code, Colorize.CSharp);
                else if (lang == "vb" || lang == "vbnet" || lang == "vb.net") lines = Colorizer.Colorize("vb", code, Colorize.VB);
                else if (lang == "" || lang == "xml") lines = Colorizer.Colorize("plain", code, Colorize.PlainText);
                else if (lang == "antlr") lines = Colorizer.Colorize("antlr", code, Antlr.ColorizeAntlr);
                else { Report.Error("MD09", $"unrecognized language {lang}"); lines = Colorize.PlainText(code); }
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

            else if (md.IsTableBlock)
            {
                var mdt = md as MarkdownParagraph.TableBlock;
                var header = mdt.headers.Option();
                var align = mdt.alignments;
                var rows = mdt.rows;
                var table = new Table();
                if (header == null) Report.Error("MD10", "Github requires all tables to have header rows");
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
                            if (p == null) { cell.Append(pars[ip]); continue; }
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
                Report.Error("MD11", $"Unrecognized markdown element {md.GetType().Name}");
                yield return new Paragraph(new Run(new Text($"[{md.GetType().Name}]")));
            }
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
                if (isOrdered.ContainsKey(level) && isOrdered[level] != isItemOrdered) Report.Error("MD12", "List can't mix ordered and unordered items at same level");
                isOrdered[level] = isItemOrdered;
                if (level > 3) Report.Error("MD13", "Can't have more than 4 levels in a list");
            }
            return flat;
        }

        IEnumerable<FlatItem> FlattenList(MarkdownParagraph.ListBlock md, int level)
        {
            var isOrdered = md.kind.IsOrdered;
            var items = md.items;
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
                        yield return new FlatItem(buglevel, wasFirstParagraph, isOrdered, mdp1);
                    }
                    else if (mdp.IsQuotedBlock || mdp.IsCodeBlock)
                    {
                        yield return new FlatItem(level, false, isOrdered, mdp);
                    }
                    else if (mdp.IsListBlock)
                    {
                        foreach (var subitem in FlattenList(mdp as MarkdownParagraph.ListBlock, level + 1)) yield return subitem;
                    }
                    else if (mdp.IsTableBlock)
                    {
                        yield return new FlatItem(level, false, isOrdered, mdp);
                    }
                    else
                    {
                        Report.Error("MD14", $"nothing fancy allowed in lists - specifically not '{mdp.GetType().Name}'");
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
            Report.CurrentSpan = md;
            if (md.IsLiteral)
            {
                var mdl = md as MarkdownSpan.Literal;
                var s = MarkdownUtilities.UnescapeLiteral(mdl);
                foreach (var r in Literal2Elements(s, nestedSpan)) yield return r;
            }

            else if (md.IsStrong || md.IsEmphasis)
            {
                IEnumerable<MarkdownSpan> spans = (md.IsStrong ? (md as MarkdownSpan.Strong).body : (md as MarkdownSpan.Emphasis).body);

                // Workaround for https://github.com/tpetricek/FSharp.formatting/issues/389 - the markdown parser
                // turns *this_is_it* into a nested Emphasis["this", Emphasis["is"], "it"] instead of Emphasis["this_is_it"]
                // What we'll do is preprocess it into Emphasis["this_is_it"]
                if (md.IsEmphasis)
                {
                    var spans2 = spans.Select(s =>
                    {
                        var _ = "";
                        if (s.IsEmphasis) { s = (s as MarkdownSpan.Emphasis).body.Single(); _ = "_"; }
                        if (s.IsLiteral) return _ + (s as MarkdownSpan.Literal).text + _;
                        Report.Error("MD15", $"something odd inside emphasis '{s.GetType().Name}' - only allowed emphasis and literal"); return "";
                    });
                    spans = new List<MarkdownSpan>() { MarkdownSpan.NewLiteral(string.Join("", spans2), FSharpOption<MarkdownRange>.None) };
                }

                // Convention is that ***term*** is used to define a term.
                // That's parsed as Strong, which contains Emphasis, which contains one Literal
                string literal = null;
                TermRef termdef = null;
                if (!nestedSpan && md.IsStrong && spans.Count() == 1 && spans.First().IsEmphasis)
                {
                    var spans2 = (spans.First() as MarkdownSpan.Emphasis).body;
                    if (spans2.Count() == 1 && spans2.First().IsLiteral)
                    {
                        literal = (spans2.First() as MarkdownSpan.Literal).text;
                        termdef = new TermRef(literal, Report.Location);
                        if (Terms.ContainsKey(literal))
                        {
                            var def = Terms[literal];
                            Report.Warning("MD16", $"Term '{literal}' defined a second time");
                            Report.Warning("MD16b", $"Here was the previous definition of term '{literal}'", def.Loc);
                        }
                        else { Terms.Add(literal, termdef); TermKeys.Clear(); }
                    }
                }

                // Convention inside our specs is that emphasis only ever contains literals,
                // either to emphasis some human-text or to refer to an ANTLR-production
                ProductionRef prodref = null;
                if (!nestedSpan && md.IsEmphasis && (spans.Count() != 1 || !spans.First().IsLiteral)) Report.Error("MD17", $"something odd inside emphasis");
                if (!nestedSpan && md.IsEmphasis && spans.Count() == 1 && spans.First().IsLiteral)
                {
                    literal = (spans.First() as MarkdownSpan.Literal).text;
                    prodref = Productions.FirstOrDefault(pr => pr.ProductionNames.Contains(literal));
                    Italics.Add(new ItalicUse(literal, prodref != null ? ItalicUse.ItalicUseKind.Production : ItalicUse.ItalicUseKind.Italic, Report.Location));
                }

                if (prodref != null)
                {
                    var props = new RunProperties(new Color { Val = "6A5ACD" }, new Underline { Val = UnderlineValues.Single });
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
                var code = mdi.code;

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
                    spans = mddl.body;
                    url = mddl.link;
                    alt = mddl.title.Option();
                }
                else
                {
                    var mdil = md as MarkdownSpan.IndirectLink;
                    var original = mdil.original;
                    var id = mdil.key;
                    spans = mdil.body;
                    if (Mddoc.DefinedLinks.ContainsKey(id))
                    {
                        url = Mddoc.DefinedLinks[id].Item1;
                        alt = Mddoc.DefinedLinks[id].Item2.Option();
                    }
                }

                var anchor = "";
                if (spans.Count() == 1 && spans.First().IsLiteral) anchor = MarkdownUtilities.UnescapeLiteral(spans.First() as MarkdownSpan.Literal);
                else if (spans.Count() == 1 && spans.First().IsInlineCode) anchor = (spans.First() as MarkdownSpan.InlineCode).code;
                else { Report.Error("MD18", $"Link anchor must be Literal or InlineCode, not '{md.GetType().Name}'"); yield break; }

                if (Sections.ContainsKey(url))
                {
                    var section = Sections[url];
                    if (anchor != section.Title) Report.Warning("MD19", $"Mismatch: link anchor is '{anchor}', should be '{section.Title}'");
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
                    Report.Error("MD20", $"Hyperlink url '{url}' unrecognized - not a recognized heading, and not http");
                }
            }

            else if (md.IsHardLineBreak)
            {
                // I've only ever seen this arise from dodgy markdown parsing, so I'll ignore it...
            }

            else
            {
                Report.Error("MD20", $"Unrecognized markdown element {md.GetType().Name}");
                yield return new Run(new Text($"[{md.GetType().Name}]"));
            }
        }

        static List<int> needleCounts = new List<int>(200);

        static IEnumerable<Needle> FindNeedles(IEnumerable<string> needles0, string haystack)
        {
            IList<string> needles = (needles0 as IList<string>) ?? new List<string>(needles0);
            for (int i = 0; i < Math.Min(needleCounts.Count, needles.Count); i++)
            {
                needleCounts[i] = 0;
            }

            while (needleCounts.Count < needles.Count)
            {
                needleCounts.Add(0);
            }
            
            var xcount = 0;
            for (int ic = 0; ic < haystack.Length; ic++)
            {
                var c = haystack[ic];
                xcount++;
                for (int i = 0; i < needles.Count; i++)
                {
                    if (needles[i][needleCounts[i]] == c)
                    {
                        needleCounts[i]++;
                        if (needleCounts[i] == needles[i].Length)
                        {
                            if (xcount > needleCounts[i])
                            {
                                yield return new Needle(-1, ic + 1 - xcount, xcount - needleCounts[i]);
                            }
                            yield return new Needle(i, ic + 1 - needleCounts[i], needleCounts[i]);
                            xcount = 0;
                            for (int j = 0; j < needles.Count; j++)
                            {
                                needleCounts[j] = 0;
                            }
                            break;
                        }
                    }
                    else
                    {
                        needleCounts[i] = 0;
                    }
                }
            }
            if (xcount > 0)
            {
                yield return new Needle(-1, haystack.Length - xcount, xcount);
            }
        }


        IEnumerable<OpenXmlElement> Literal2Elements(string literal, bool isNested)
        {
            if (isNested || Terms.Count == 0)
            {
                yield return new Run(new Text(literal) { Space = SpaceProcessingModeValues.Preserve });
                yield break;
            }

            if (TermKeys.Count == 0) TermKeys.AddRange(Terms.Keys);

            foreach (var needle in FindNeedles(TermKeys, literal))
            {
                var s = literal.Substring(needle.Start, needle.Length);
                if (needle.NeedleId == -1)
                {
                    yield return new Run(new Text(s) { Space = SpaceProcessingModeValues.Preserve });
                    continue;
                }
                var termref = Terms[s];
                Italics.Add(new ItalicUse(s, ItalicUse.ItalicUseKind.Term, Report.Location));
                var props = new RunProperties(new Underline { Val = UnderlineValues.Dotted, Color = "4BACC6" });
                var run = new Run(new Text(s) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props };
                var link = new Hyperlink(run) { Anchor = termref.BookmarkName };
                yield return link;
            }

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
            var spans = p.body;
            if (spans.Count() == 0 || !spans[0].IsLiteral) return level;
            var literal = spans[0] as MarkdownSpan.Literal;
            if (!literal.text.StartsWith("ceci-n'est-pas-une-indent")) return level;
            //
            var literal2 = MarkdownSpan.NewLiteral(literal.text.Substring(25), FSharpOption<MarkdownRange>.None);
            var spans2 = Microsoft.FSharp.Collections.FSharpList<MarkdownSpan>.Cons(literal2, spans.Tail);
            var p2 = MarkdownParagraph.NewParagraph(spans2, FSharpOption<MarkdownRange>.None);
            mdp = p2;
            return 0;
        }
    }
}
