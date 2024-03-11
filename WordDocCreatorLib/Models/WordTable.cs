namespace WordDocCreatorLib.Models
{
    /// <summary>
    /// This defines the table data and it's styling properties 
    /// to apply to the word document table. 
    /// </summary>
    public class WordTable
    {
        private int last = -1;

        public WordTable(int rows, int columns) 
        { 
            Data = new string[rows, columns];
            WordTableStyleName = new WordTableStyleName();
            WordFont = new WordFont();
        }

        public string[,] Data { get; private set; }

        /// <summary>
        /// This is the name of the table style that is pre-defined in the Word document. 
        /// </summary>
        public WordTableStyleName WordTableStyleName { get; set; }

        /// <summary>
        /// This contains the font properties of the font defined in Word.
        /// </summary>
        public WordFont WordFont { get; set; }

        public WordTable AddData(string[] row)
        {
            if(row == null || row.Length == 0)
            {
                throw new Exception("Row data must have a value and non-zero length.");
            }

            last++;
            
            for(int i = 0; i < row.Length; i++)
            {
                Data[last, i] = row[i];
            }

            return this;
        }

        public override string ToString()
        {
            return WordTableStyleName.StyleName;
        }
    }
}