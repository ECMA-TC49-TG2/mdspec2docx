using FSharp.Markdown;

namespace MarkdownConverter.Spec
{
    // TODO: Potentially move this to the root, and refactor somewhat.

    /// <summary>
    /// Diagnostic reporter
    /// </summary>
    internal class Reporter
    {
        public SourceLocation Location = new SourceLocation(null, null, null, null);

        public string CurrentFile
        {
            get => Location.file;
            set => Location = new SourceLocation(value, null, null, null);
        }

        public SectionRef CurrentSection
        {
            get => Location.section;
            set => Location = new SourceLocation(CurrentFile, value, CurrentParagraph, null);
        }

        public MarkdownParagraph CurrentParagraph
        {
            get => Location.paragraph;
            set => Location = new SourceLocation(CurrentFile, CurrentSection, value, null);
        }

        public MarkdownSpan CurrentSpan
        {
            get => Location.span;
            set => Location = new SourceLocation(CurrentFile, CurrentSection, CurrentParagraph, value);
        }

        public void Error(string code, string msg, SourceLocation loc = null) => Program.Report(code, "ERROR", msg, loc?.loc ?? Location.loc);

        public void Warning(string code, string msg, SourceLocation loc = null) => Program.Report(code, "WARNING", msg, loc?.loc ?? Location.loc);

        public void Log(string msg) { }

    }
}
