namespace WordDocCreatorLib.Core.Models
{
    public class WordText
    {
        public WordText(string text, string? hyperlink = null, bool doNotUnderlineHyperlink = false, WordFont? wordFont = null)
        {
            Text = text;
            Hyperlink = hyperlink;
            DoNotUnderlineHyperlink = doNotUnderlineHyperlink;
            WordFont = wordFont;
        }

        public string Text { get; set; }
        public string? Hyperlink { get; set; }
        public bool DoNotUnderlineHyperlink { get; set; }
        public WordFont? WordFont { get; set; }

        public override string ToString()
        {
            return Text;
        }

    }
}