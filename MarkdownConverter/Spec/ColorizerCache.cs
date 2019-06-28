using CSharp2Colorized;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MarkdownConverter.Spec
{
    class ColorizerCache
    {
        class CacheEntry
        {
            public bool isUsed;
            public List<ColorizedLine> clines;
        }

        private Dictionary<string, CacheEntry> cache = new Dictionary<string, CacheEntry>();
        private bool hasNew;
        private string fn;
        private string tempdir;

        public ColorizerCache(string tempdir)
        {
            this.tempdir = tempdir;
            if (tempdir == null) return;
            fn = tempdir + "\\colorized.cache.txt";
            if (!File.Exists(fn)) return;

            using (var s = new StreamReader(fn))
            {
                while (true)
                {
                    var code = s.ReadLine();
                    var clines = s.ReadLine();
                    if (code == "EOF" && clines == null) return;
                    if (code == null || clines == null) { cache.Clear(); return; }
                    cache[DeserializeCode(code)] = new CacheEntry { clines = DeserializeColor(clines) };
                }
            }
        }

        public void Close()
        {
            if (fn == null) return;
            if (!hasNew && cache.All(kv => kv.Value.isUsed)) return;
            using (var s = new StreamWriter(fn))
            {
                foreach (var kv in cache)
                {
                    if (!kv.Value.isUsed) continue;
                    s.WriteLine(SerializeCode(kv.Key));
                    s.WriteLine(SerializeColor(kv.Value.clines));
                }
                s.WriteLine("EOF");
            }
        }

        public IEnumerable<ColorizedLine> Colorize(string lang, string code, Func<string, IEnumerable<ColorizedLine>> generator)
        {
            string key = $"{lang}:{code}";
            CacheEntry e;
            if (cache.TryGetValue(key, out e))
            {
                e.isUsed = true;
                return e.clines;
            }
            hasNew = true;
            e = new CacheEntry { isUsed = true, clines = generator(code).ToList() };
            cache.Add(key, e);
            return e.clines;
        }

        static string SerializeCode(string code)
        {
            return code.Replace("\\", "\\\\").Replace("\r", "\\r").Replace("\n", "\\n");
        }

        static string DeserializeCode(string src)
        {
            return src.Replace("\\r", "\r").Replace("\\n", "\n").Replace("\\\\", "\\");
        }

        static string SerializeColor(IEnumerable<ColorizedLine> lines)
        {
            var sb = new StringBuilder();
            var needComma = false;
            foreach (var line in lines)
            {
                if (needComma) sb.Append(",");
                foreach (var word in line.Words)
                {
                    sb.Append(word.IsItalic ? "i#" : "r#");
                    sb.Append($"{word.Red:X2}{word.Green:X2}{word.Blue:X2}");
                    sb.Append(word.Text.Replace("\\", "\\\\").Replace(",", "\\c").Replace("\n", "\\n").Replace("\r", "\\r"));
                    sb.Append(",");
                }
                needComma = true;
            }
            return sb.ToString();
        }

        static List<ColorizedLine> DeserializeColor(string src)
        {
            var atoms = src.Split(',');
            var lines = new List<ColorizedLine>();
            var line = new ColorizedLine();
            foreach (var atom in atoms)
            {
                if (atom == "") { lines.Add(line); line = new ColorizedLine(); continue; }
                // i#RRGGBBt
                var w = new ColorizedWord();
                if (atom[0] == 'i') w.IsItalic = true;
                var col = Convert.ToUInt32(atom.Substring(2, 6), 16);
                w.Red = (int) ((col >> 16) & 255);
                w.Green = (int) ((col >> 8) & 255);
                w.Blue = (int) ((col >> 0) & 255);
                w.Text = atom.Substring(8).Replace("\\r", "\r").Replace("\\n", "\n").Replace("\\c", ",").Replace("\\\\", "\\");
                line.Words.Add(w);
            }
            return lines;
        }
    }
}
