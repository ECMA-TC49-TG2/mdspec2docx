namespace MarkdownConverter.Spec
{
    internal class ItalicUse
    {
        public string Literal;
        public ItalicUseKind Kind;
        public SourceLocation Loc;
        public enum ItalicUseKind { Production, Italic, Term };

        public ItalicUse(string literal, ItalicUseKind kind, SourceLocation loc)
        {
            Literal = literal;
            Kind = kind;
            Loc = loc;
        }
    }
}
