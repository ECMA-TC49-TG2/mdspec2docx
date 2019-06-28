using System;
using System.Collections.Generic;
using System.Text;

namespace MarkdownConverter.Spec
{
    internal class Span
    {
        public int start, length;
        public Span(int start, int length) { this.start = start; this.length = length; }
        public int end => start + length;
    }
}
