namespace WordDocCreatorLib.Core.Models
{
    public class WordFont
    {
        public WordFont()
        {
            WordFontName = new WordFontName();
            Size = 10;
        }

        public WordFont(WordFontName wordFontName, float size)
        {
            WordFontName = wordFontName;
            Size = size;
        }

        /// <summary>
        /// This is the name of the font that is pre-defined in the Word document. 
        /// </summary>
        public WordFontName WordFontName { get; set; }

        public float Size { get; set; }

        public override string ToString()
        {
            return WordFontName.Name;
        }

    }
}