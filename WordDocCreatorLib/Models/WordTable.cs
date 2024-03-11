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
        }

        public string[,] Data { get; private set; }

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
    }
}