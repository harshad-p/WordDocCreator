﻿using WordDocCreatorLib;

namespace WordDocCreatorApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("***Recipe Binder Doc. Creator***");

            var directory = "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Automated";
            var fileName = "01 Shengdanyachi Bhaaji - Sarika Nilatkar - A5";
            var templatePath = "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Templates\\01 A5 - Kohinoor.dotx";

            var wordDocCreator = new WordDocCreator(templatePath);

            wordDocCreator.FillShapeWithImage("Recipe_Image", "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\01 Veg\\01 Shengdanyachi Bhaaji\\IMG_0161.HEIC");
            wordDocCreator.UpdateBookmarkedText("Recipe_Title", "Shengdanyachi Bhaaji");
            wordDocCreator.UpdateBookmarkedText("Recipe_Author", "Sarika Nilatkar");
            wordDocCreator.UpdateBookmarkedText("Procedure", "Shengdane bhajun tyancha kut karun ghyaycha.");
            Console.WriteLine(wordDocCreator.SaveAs(directory, fileName, SaveAsDocumentType.DOCX));

            wordDocCreator.Dispose();
        }
    }
}
