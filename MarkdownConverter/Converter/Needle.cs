namespace MarkdownConverter.Converter
{
    internal struct Needle
    {
        public int NeedleId; // or -1 if this was a non-matching span
        public int Start { get; }
        public int Length { get; }

        public Needle(int needleId, int start, int length)
        {
            NeedleId = needleId;
            Start = start;
            Length = length;
        }
    }
}
