using WordDocCreatorLib.Models;

namespace WordDocCreatorApp
{
    public class WordTemplateInput
    {
        public string Directory { get; set; }

        public string FileName { get; set; }

        public string TemplatePath { get; set; }

        public IDictionary<string, WordTable> WordTables { get; set; }

        public IDictionary<string, string> Images { get; set; }

        public IDictionary<string, Tuple<string, string>> Texts { get; set; }

    }
}