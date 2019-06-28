using MarkdownConverter.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarkdownConverter.Spec
{
    internal class ProductionRef
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
}
