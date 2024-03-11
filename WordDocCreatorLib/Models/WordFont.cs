﻿namespace WordDocCreatorLib.Models
{
    public class WordFont
    {
        public WordFont() 
        {
            WordFontName = new WordFontName();
            Size = 10;
        }

        /// <summary>
        /// This is the name of the font that is pre-defined in the Word document. 
        /// </summary>
        public WordFontName WordFontName { get; set; }

        public int Size { get; set; }

        public override string ToString()
        {
            return WordFontName.Name;
        }

    }
}