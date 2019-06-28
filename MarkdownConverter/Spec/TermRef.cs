namespace MarkdownConverter.Spec
{
    internal class TermRef
    {
        public string Term;
        public string BookmarkName;
        public SourceLocation Loc;
        public static int count = 1;

        public TermRef(string term, SourceLocation loc)
        {
            Term = term;
            Loc = loc;
            BookmarkName = $"_Trm{count:00000}"; count++;
        }
    }
}
