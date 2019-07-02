using FSharp.Markdown;

namespace MarkdownConverter.Spec
{
    // TODO: Potentially move this to the root, and refactor somewhat.

    /// <summary>
    /// Diagnostic reporter
    /// </summary>
    internal class Reporter
    {
        public SourceLocation Location { get; set; } = new SourceLocation(null, null, null, null);

        public Reporter(string filename)
        {
            Location = new SourceLocation(filename, null, null, null);
        }

        public string CurrentFile => Location.File;

        public SectionRef CurrentSection
        {
            get => Location.Section;
            set => Location = new SourceLocation(CurrentFile, value, CurrentParagraph, null);
        }

        public MarkdownParagraph CurrentParagraph
        {
            get => Location.Paragraph;
            set => Location = new SourceLocation(CurrentFile, CurrentSection, value, null);
        }

        public MarkdownSpan CurrentSpan
        {
            get => Location.Span;
            set => Location = new SourceLocation(CurrentFile, CurrentSection, CurrentParagraph, value);
        }

        public void Error(string code, string msg, SourceLocation loc = null) => Program.Report(code, "ERROR", msg, loc?.loc ?? Location.loc);

        public void Warning(string code, string msg, SourceLocation loc = null) => Program.Report(code, "WARNING", msg, loc?.loc ?? Location.loc);

        public void Log(string msg) { }

    }
}
