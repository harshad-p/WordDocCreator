using InputProviderLib;
using WordDocCreatorLib;
using WordDocCreatorLib.Core.Models;

namespace WordDocCreatorApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("***Word Doc. Creator***");

            IWordTemplateInputProvider wordTemplateInputProvider = new SampleRecipeBinderWordTemplateInputProvider();

            var wordTemplateInputs = wordTemplateInputProvider.GetTemplateInputs();

            // create a single instance of Word instead of one for every document.
            using var wordDocCreator = new WordDocCreator();

            foreach (var wordTemplateInput in wordTemplateInputs)
            {
                // load the template again as the bookmarks get deleted after it is replaced with content
                wordDocCreator.LoadTemplate(wordTemplateInput.TemplateFilePath);

                foreach (var wordDocumentInput in wordTemplateInput.WordDocumentInputs)
                {
                    FillImages(wordDocCreator, wordDocumentInput.Images);
                    FillTables(wordDocCreator, wordDocumentInput.WordTables);
                    FillTexts(wordDocCreator, wordDocumentInput.Texts);

                    // Need to save the docx file first...
                    var fileSavePath = wordDocCreator.SaveAs(wordDocumentInput.SaveDirectory, wordDocumentInput.FileName, SaveAsDocumentType.DOCX);
                    // ... to be exported as a pdf first.
                    wordDocCreator.SaveAs(wordDocumentInput.SaveDirectory, wordDocumentInput.FileName, SaveAsDocumentType.PDF);
                    Console.WriteLine(fileSavePath);
                }
            }
        }

        private static void FillTexts(WordDocCreator wordDocCreator, IDictionary<string, WordText> texts)
        {
            foreach (var (templateFile, text) in texts)
            {
                wordDocCreator.UpdateBookmarkedText(templateFile, text.Text, text.Hyperlink, text.DoNotUnderlineHyperlink, text.WordFont);
            }
        }

        private static void FillTables(WordDocCreator wordDocCreator, IDictionary<string, WordTable> wordTables)
        {
            foreach (var wordTable in wordTables)
            {
                wordDocCreator.InsertTable(wordTable.Key, wordTable.Value);
            }
        }

        private static void FillImages(WordDocCreator wordDocCreator, IDictionary<string, string> images)
        {
            foreach (var image in images)
            {
                wordDocCreator.FillShapeWithImage(image.Key, image.Value);
            }
        }

    }
}
