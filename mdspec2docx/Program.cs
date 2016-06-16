﻿// TODO:
//
// * Something goofy is going on with the TOC in the Word document. The field codes are broken, so it
//   isn't recognized as a field. And if I edit the field in Word (e.g. to add page numbers) then all
//   within-spec section links get broken. Also, md2docx should only generate level 1+2 TOC headings.
// * Make a better "report error" thing, and distinguish fatal from non-fatal (e.g. duplicate section title)

using Grammar2Html;
using Microsoft.FSharp.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

static class Program
{
    static int Main(string[] args)
    {
        var ifn = (args.Length >= 2 ? args[1] : "readme.md");

        if (!File.Exists(ifn) || !File.Exists("template.docx") || Directory.GetFiles(".", "*.g4").Length > 1)
        {
            Console.Error.WriteLine("md2docx <filename>.md -- converts it to '<filename>.docx', based on 'template.docx'");
            Console.Error.WriteLine();
            Console.Error.WriteLine("If no file is specified:");
            Console.Error.WriteLine("    it looks for readme.md instead");
            Console.Error.WriteLine("If input file has a list with links of the form `* [Link](subfile.md)`:");
            Console.Error.WriteLine("   it converts the listed subfiles instead of <filename>.md");
            Console.Error.WriteLine("If the current directory contains one <grammar>.g4 file:");
            Console.Error.WriteLine("   it verifies all ```antlr blocks correspond, and also generates <grammar>.html");
            Console.Error.WriteLine("If 'template.docx' contains a Table of Contents:");
            Console.Error.WriteLine("   it replaces it with one based on the markdown (but page numbers aren't supported)");
            return 1;
        }

        // Read input file. If it contains a load of linked filenames, then read them instead.
        var readme = FSharp.Markdown.Markdown.Parse(File.ReadAllText(ifn));
        var files = (from list in readme.Paragraphs.OfType<FSharp.Markdown.MarkdownParagraph.ListBlock>()
                     let items = list.Item2
                     from par in items
                     from spanpar in par.OfType<FSharp.Markdown.MarkdownParagraph.Span>()
                     let spans = spanpar.Item
                     from link in spans.OfType<FSharp.Markdown.MarkdownSpan.DirectLink>()
                     let url = link.Item2.Item1
                     where url.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase)
                     select url).ToList().Distinct();
        if (files.Count() == 0) files = new[] { ifn };
        var md = MarkdownSpec.ReadFiles(files);


        // Now md.Gramar contains the grammar as extracted out of the *.md files, and moreover has
        // correct references to within the spec. We'll check that it has the same productions as
        // in the corresponding ANTLR file
        var antlrfn = Directory.GetFiles(".", "*.g4").FirstOrDefault();
        if (antlrfn != null)
        {
            var htmlfn = Path.ChangeExtension(antlrfn, ".html");
            var grammar = Antlr.ReadFile(antlrfn);
            foreach (var diff in CompareGrammars(grammar, md.Grammar))
            {
                if (diff.authority == null) Console.WriteLine($"mdspec2docx: error MD021: markdown has superfluous production '{diff.productionName}'");
                else if (diff.copy == null) Console.WriteLine($"mdspec2docx: error MD022: markdown lacks production '{diff.productionName}'");
                else
                {
                    Console.WriteLine($"mdspec2docx: error MD023: production '{diff.productionName}' differs between markdown and antlr.g4");
                    Console.WriteLine("mdspec2docx: error MD023b: antlr.g4 says " + diff.authority.Replace("\r", "\\r").Replace("\n", "\\n"));
                    Console.WriteLine("mdspec2docx: error MD023c: markdown says " + diff.copy.Replace("\r", "\\r").Replace("\n", "\\n"));
                }
            }
            foreach (var p in grammar.Productions)
            {
                p.Link = md.Grammar.Productions.FirstOrDefault(mdp => mdp?.ProductionName == p.ProductionName)?.Link;
                p.LinkName = md.Grammar.Productions.FirstOrDefault(mdp => mdp?.ProductionName == p.ProductionName)?.LinkName;
            }

            File.WriteAllText(htmlfn, grammar.ToHtml(), Encoding.UTF8);
            Process.Start(htmlfn);
        }

        // Generate the Specification.docx file
        var fn = PickUniqueFilename(Path.ChangeExtension(ifn, ".docx"));
        md.WriteFile("template.docx", fn);
        Process.Start(fn);
        return 0;
    }

    static string PickUniqueFilename(string suggestion)
    {
        var dir = Path.GetFileNameWithoutExtension(suggestion);
        var ext = Path.GetExtension(suggestion);

        for (int ifn=0; true; ifn++)
        { 
            var fn = dir + (ifn == 0 ? "" : ifn.ToString()) + ext;
            if (!File.Exists(fn)) return fn;
            try { File.Delete(fn); return fn; } catch (Exception) { }
            ifn++;
        }
    }

    class ProductionDifference
    {
        public string productionName;
        public string authority, copy;
    }

    static IEnumerable<ProductionDifference> CompareGrammars(Grammar authority, Grammar copy)
    {
        Func<Grammar, Dictionary<string, Production>> ToDictionary;
        ToDictionary = g =>
        {
            var d = new Dictionary<string, Production>();
            foreach (var pp in g.Productions) if (pp.ProductionName != null) d[pp.ProductionName] = pp;
            return d;
        };
        var dauthority = ToDictionary(authority);
        var dcopy = ToDictionary(copy);

        foreach (var p in dauthority.Keys)
        {
            if (!dcopy.ContainsKey(p)) continue;
            Production pauthority0 = dauthority[p], pcopy0 = dcopy[p];
            string pauthority = Antlr.ToString(pauthority0), pcopy = Antlr.ToString(pcopy0);
            if (pauthority == pcopy) continue;
            yield return new ProductionDifference { productionName = p, authority = pauthority, copy = pcopy };
        }

        foreach (var p in dauthority.Keys)
        {
            if (p == "start") continue;
            if (!dcopy.ContainsKey(p)) yield return new ProductionDifference { productionName = p, authority = "<defined>", copy = null };
        }
        foreach (var p in dcopy.Keys)
        {
            if (p == "start") continue;
            if (!dauthority.ContainsKey(p)) yield return new ProductionDifference { productionName = p, authority = null, copy = "<defined>" };
        }
    }


    public static T Option<T>(this FSharpOption<T> o) where T : class
    {
        if (FSharpOption<T>.GetTag(o) == FSharpOption<T>.Tags.None) return null;
        return o.Value;
    }

}
