using WordDocCreatorLib;
using WordDocCreatorLib.Models;

namespace WordDocCreatorApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("***Word Doc. Creator***");

            var wordTable = new WordTable(3, 2);
            wordTable.AddData(["Shengdane", "Ardha Kilo"]).AddData(["Oil", "2 tbl. spoons"]).AddData(["Salt", "Chavinusar"]);

            var wordTemplateInput = new WordTemplateInput
            {
                Directory = "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Automated",
                FileName = "01 Shengdanyachi Bhaaji - Sarika Nilatkar - A5",
                TemplatePath = "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Templates\\01 A5 - Kohinoor.dotx",
                WordTables = new Dictionary<string, WordTable>
                {
                    { "Ingredients_Table", wordTable }
                },
                Images = new Dictionary<string, string>
                {
                    { "Recipe_Image", "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\01 Veg\\01 Shengdanyachi Bhaaji\\IMG_0161.HEIC" }
                },
                Texts = new Dictionary<string, Tuple<string, string>>
                {
                    { "Recipe_Title", Tuple.Create<string, string>("Shengdanyachi Bhaaji", null) },
                    { "Recipe_Author", Tuple.Create("Sarika Nilatkar" , "www.google.com")},
                    { "Procedure", Tuple.Create<string, string>("Shengdane bhajun tyancha kut karun ghyaycha." , null)},
                    { "Recipe_Is_Ready", Tuple.Create<string, string>("Shengdanyachi Bhaaji Tayyar" , null) }
                }
            };

            var wordDocCreator = new WordDocCreator(wordTemplateInput.TemplatePath);

            FillImages(wordDocCreator, wordTemplateInput.Images);
            FillTables(wordDocCreator, wordTemplateInput.WordTables);
            FillTexts(wordDocCreator, wordTemplateInput.Texts);

            var fileSavePath = wordDocCreator.SaveAs(wordTemplateInput.Directory, wordTemplateInput.FileName, SaveAsDocumentType.DOCX);
            Console.WriteLine(fileSavePath);

            wordDocCreator.Dispose();
        }

        private static void FillTexts(WordDocCreator wordDocCreator, IDictionary<string, Tuple<string, string>> texts)
        {
            foreach (var text in texts)
            {
                wordDocCreator.UpdateBookmarkedText(text.Key, text.Value.Item1, text.Value.Item2);
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
