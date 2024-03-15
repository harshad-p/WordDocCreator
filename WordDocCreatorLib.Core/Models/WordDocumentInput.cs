namespace WordDocCreatorLib.Core.Models
{
    public class WordDocumentInput
    {
        /// <summary>
        /// The directory in which to place the created word document. 
        /// </summary>
        public string SaveDirectory { get; set; }

        /// <summary>
        /// The name (without extension) of the file to save the word document as. 
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Any table data that needs to be added in the bookmarked region  
        /// denoted by the dictionary key. 
        /// </summary>
        public IDictionary<string, WordTable> WordTables { get; set; }

        /// <summary>
        /// Paths to any images that need to be inserted into the bookmarked region  
        /// denoted by the dictionary key. 
        /// </summary>
        public IDictionary<string, string> Images { get; set; }

        /// <summary>
        /// Any text that needs to be added in the bookmarked region 
        /// denoted by the dictionary key. 
        /// </summary>
        public IDictionary<string, Tuple<string, string>> Texts { get; set; }

    }
}