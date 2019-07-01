using FSharp.Markdown;

namespace MarkdownConverter.Converter
{
    internal class FlatItem
    {
        public int Level;
        public bool HasBullet;
        public bool IsBulletOrdered;
        public MarkdownParagraph Paragraph;
    }
}
