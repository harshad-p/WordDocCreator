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
}