﻿using WordDocCreatorLib;
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

                        var fileSavePath = wordDocCreator.SaveAs(wordDocumentInput.Directory, wordDocumentInput.FileName, SaveAsDocumentType.DOCX);
                        Console.WriteLine(fileSavePath);
                    }
                }
            }
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
