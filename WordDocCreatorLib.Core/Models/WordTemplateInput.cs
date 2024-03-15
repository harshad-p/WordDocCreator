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

        /// <summary>
        /// The path at which the word template file is located. 
        /// </summary>
        public string TemplateFilePath { get; set; }

        /// <summary>
        /// The inputs to create a word document out of the template file. 
        /// </summary>
        public IEnumerable<WordDocumentInput> WordDocumentInputs { get; set; }

    }
}