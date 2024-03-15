namespace WordDocCreatorLib.Core.Models
{
    public class WordDocumentInput
    {
        public string Directory { get; set; }

        public string FileName { get; set; }

        public IDictionary<string, WordTable> WordTables { get; set; }

        public IDictionary<string, string> Images { get; set; }

        public IDictionary<string, Tuple<string, string>> Texts { get; set; }

    }
}