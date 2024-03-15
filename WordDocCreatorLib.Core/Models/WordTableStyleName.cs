namespace WordDocCreatorLib.Core.Models
{
    public class WordTableStyleName
    {
        public WordTableStyleName()
        {
            StyleName = TableGrid.StyleName;
        }

        public WordTableStyleName(string styleName)
        {
            StyleName = styleName;
        }

        /// <summary>
        /// This is the name of the table style that is pre-defined in the Word document. 
        /// </summary>
        public string StyleName { get; set; }

        public static readonly WordTableStyleName TableGrid = new WordTableStyleName("Table Grid");
        public static readonly WordTableStyleName TableGridLight = new WordTableStyleName("Table Grid Light");
        public static readonly WordTableStyleName PlainTable1 = new WordTableStyleName("Plain Table 1");
        public static readonly WordTableStyleName PlainTable2 = new WordTableStyleName("Plain Table 2");
        public static readonly WordTableStyleName PlainTable3 = new WordTableStyleName("Plain Table 3");

        public override string ToString()
        {
            return StyleName;
        }

    }
}