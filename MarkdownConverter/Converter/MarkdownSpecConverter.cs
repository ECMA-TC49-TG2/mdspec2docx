using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdownConverter.Spec;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace MarkdownConverter.Converter
{
    internal static class MarkdownSpecConverter
    {
        public static void ConvertToWord(MarkdownSpec spec, string templateFile, string outputFile, string tempDir)
        {
            using (var templateDoc = WordprocessingDocument.Open(templateFile, false))
            using (var resultDoc = WordprocessingDocument.Create(outputFile, WordprocessingDocumentType.Document))
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
                    for (int i = 0; i < spec.Sections.Count; i++)
                    {
                        var section = spec.Sections[i];
                        if (section.Level > 2) continue;
                        var p = new Paragraph();
                        if (i == 0) p.AppendChild(tocRunFirst);
                        p.AppendChild(new Hyperlink(new Run(new Text(section.Number + " " + section.Title))) { Anchor = section.BookmarkName });
                        if (i == spec.Sections.Count - 1) p.AppendChild(tocRunLast);
                        p.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = $"TOC{section.Level}" });
                        body.InsertBefore(p, afterToc);
                    }
                    if (tocSec != null) body.InsertBefore(tocSec, afterToc);
                }

                var maxBookmarkId = new StrongBox<int>(1 + body.Descendants<BookmarkStart>().Max(bookmark => int.Parse(bookmark.Id)));
                var terms = new Dictionary<string, TermRef>();
                var termkeys = new List<string>();
                var italics = new List<ItalicUse>();
                var colorizer = new ColorizerCache(tempDir);
                foreach (var src in spec.Sources)
                {
                    // FIXME: Ick
                    spec.Report.CurrentFile = Path.GetFullPath(src.Item1);
                    var converter = new MarkdownSourceConverter
                    {
                        Mddoc = src.Item2,
                        Filename = Path.GetFileName(src.Item1),
                        Wdoc = resultDoc,
                        Sections = spec.Sections.ToDictionary(sr => sr.Url),
                        Productions = spec.Productions,
                        Terms = terms,
                        TermKeys = termkeys,
                        Italics = italics,
                        MaxBookmarkId = maxBookmarkId,
                        Report = spec.Report,
                        Colorizer = colorizer
                    };
                    foreach (var p in converter.Paragraphs())
                    {
                        body.AppendChild(p);
                    }
                }

                colorizer.Close();
                spec.Report.CurrentFile = null;
                spec.Report.CurrentSection = null;
                spec.Report.CurrentParagraph = null;
                spec.Report.CurrentSpan = null;

                // I wonder if there were any oddities? ...
                // Terms that were referenced before their definition?
                var termset = new HashSet<string>(terms.Keys);
                var italicset = new HashSet<string>(italics.Where(i => i.Kind == ItalicUse.ItalicUseKind.Italic).Select(i => i.Literal));
                italicset.IntersectWith(termset);
                foreach (var s in italicset)
                {
                    var use = italics.First(i => i.Literal == s);
                    var def = terms[s];
                    spec.Report.Warning("MD05", $"Term '{s}' used before definition", use.Loc);
                    spec.Report.Warning("MD05b", $"... definition location of '{s}' for previous warning", def.Loc);
                }

                // Terms that are also production names?
                var productionset = new HashSet<string>(spec.Grammar.Productions.Where(p => p.ProductionName != null).Select(p => p.ProductionName));
                productionset.IntersectWith(termset);
                foreach (var s in productionset)
                {
                    var def = terms[s];
                    spec.Report.Warning("MD06", $"Terms '{s}' is also a grammar production name", def.Loc);
                }

                // Terms that were defined but never used?
                var termrefset = new HashSet<string>(italics.Where(i => i.Kind == ItalicUse.ItalicUseKind.Term).Select(i => i.Literal));
                termset.RemoveWhere(t => termrefset.Contains(t));
                foreach (var s in termset)
                {
                    var def = terms[s];
                    spec.Report.Warning("MD07", $"Term '{s}' is defined but never used", def.Loc);
                }

                // Which single-word production-names appear in italics?
                var italicproductionset = new HashSet<string>(italics.Where(i => !i.Literal.Contains("_") && i.Kind == ItalicUse.ItalicUseKind.Production).Select(i => i.Literal));
                var italicproductions = string.Join(",", italicproductionset);

                // What are the single-word production names that don't appear in italics?
                var otherproductionset = new HashSet<string>(spec.Grammar.Productions.Where(p => p.ProductionName != null && !p.ProductionName.Contains("_") && !italicproductionset.Contains(p.ProductionName)).Select(p => p.ProductionName));
                var otherproductions = string.Join(",", otherproductionset);

                // FIXME: We're not using these last variables...
            }
        }

        private static bool FindToc(Body body, out int ifirst, out int iLast, out string instr, out Paragraph secBreak)
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
    }
}
