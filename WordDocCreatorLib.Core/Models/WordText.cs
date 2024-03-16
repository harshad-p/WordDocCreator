using System.ComponentModel.DataAnnotations;

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

        /// <summary>
        /// A required string that will be inserted into the document.
        /// </summary>
        [Required]
        public string Text { get; set; }

        /// <summary>
        /// An optional string to add to the text as a hyperlink. 
        /// </summary>
        public string? Hyperlink { get; set; }

        /// <summary>
        /// By default, Word underlines hyperlinked text. 
        /// If you prefer not to have the underline, indicate so over here. 
        /// </summary>
        public bool DoNotUnderlineHyperlink { get; set; }

        /// <summary>
        /// The default font in the template file will be applied 
        /// to the text inserted into the template. 
        /// But there are cases where devanagari text is inserted using a default font. 
        /// If you wish to provide another devanagari font, 
        /// or if you wish to override the font in the template for the text, 
        /// supply using this WordFont.
        /// </summary>
        public WordFont? WordFont { get; set; }

        public override string ToString()
        {
            return Text;
        }

    }
}