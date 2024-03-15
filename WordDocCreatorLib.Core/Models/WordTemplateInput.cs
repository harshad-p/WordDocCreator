namespace WordDocCreatorLib.Core.Models
{
    public class WordTemplateInput
    {
        public WordTemplateInput()
        {
            WordDocumentInputs = new List<WordDocumentInput>();
        }

        public WordTemplateInput(string templateFilePath)
        {
            TemplateFilePath = templateFilePath;
            WordDocumentInputs = new List<WordDocumentInput>();
        }

        public WordTemplateInput(string templateFilePath, IEnumerable<WordDocumentInput> wordDocumentInputs)
        {
            TemplateFilePath = templateFilePath;
            WordDocumentInputs = wordDocumentInputs;
        }

        public string TemplateFilePath { get; set; }
        public IEnumerable<WordDocumentInput> WordDocumentInputs { get; set; }


    }

    public class WordDocumentInput
    {
        public string Directory { get; set; }

        public string FileName { get; set; }

        public IDictionary<string, WordTable> WordTables { get; set; }

        public IDictionary<string, string> Images { get; set; }

        public IDictionary<string, Tuple<string, string>> Texts { get; set; }

    }
}