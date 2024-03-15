using WordDocCreatorLib;
using WordDocCreatorLib.Models;

namespace WordDocCreatorApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("***Word Doc. Creator***");

            var wordTemplateInputs = WordTemplateInput.GetSampleInputs();

            foreach(var wordTemplateInput in wordTemplateInputs)
            {
                // create a single instance of Word instead of one for every document.
                using var wordDocCreator = new WordDocCreator();

                foreach (var wordDocumentInputs in wordTemplateInput.WordDocumentInputs)
                {
                    // load the template again as the bookmarks get deleted after it is replaced with content
                    wordDocCreator.LoadTemplate(wordDocumentInputs.Key);

                    foreach (var wordDocumentInput in wordDocumentInputs.Value)
                    {
                        FillImages(wordDocCreator, wordDocumentInput.Images);
                        FillTables(wordDocCreator, wordDocumentInput.WordTables);
                        FillTexts(wordDocCreator, wordDocumentInput.Texts);

                        // Need to save the docx file first...
                        var fileSavePath = wordDocCreator.SaveAs(wordDocumentInput.Directory, wordDocumentInput.FileName, SaveAsDocumentType.DOCX);
                        // ... to be exported as a pdf first.
                        wordDocCreator.SaveAs(wordDocumentInput.Directory, wordDocumentInput.FileName, SaveAsDocumentType.PDF);
                        Console.WriteLine(fileSavePath);
                    }
                }
            }
        }

        private static void FillTexts(WordDocCreator wordDocCreator, IDictionary<string, Tuple<string, string>> texts, string? fontName = null, float? fontSize = null)
        {
            foreach (var text in texts)
            {
                WordFont? wordFont = null;
                if (fontName != null && fontSize != null)
                {
                    // Either provide both, or nothing
                    wordFont = new WordFont(new WordFontName(fontName), fontSize.Value);
                }
                wordDocCreator.UpdateBookmarkedText(text.Key, text.Value.Item1, text.Value.Item2, null, wordFont);
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
