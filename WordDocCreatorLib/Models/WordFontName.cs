namespace WordDocCreatorLib.Models
{
    public class WordFontName
    {
        public WordFontName()
        {
            Name = Helvetica.Name;
        }

        public WordFontName(string name)
        {
            Name = name;
        }

        /// <summary>
        /// This is the name of the font that is pre-defined in the Word document. 
        /// </summary>
        public string Name { get; set; }

        public static readonly WordFontName CalibriLight = new WordFontName("Calibri Light");
        public static readonly WordFontName Calibri = new WordFontName("Calibri");
        public static readonly WordFontName Arial = new WordFontName("Arial");
        public static readonly WordFontName Verdana = new WordFontName("Verdana");
        public static readonly WordFontName Helvetica = new WordFontName("Helvetica");

        public override string ToString()
        {
            return Name;
        }

    }
}