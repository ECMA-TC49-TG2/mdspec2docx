using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarkdownConverter.Converter
{
    internal struct Needle
    {
        public int needle; // or -1 if this was a non-matching span
        public int istart;
        public int length;
    }
}
