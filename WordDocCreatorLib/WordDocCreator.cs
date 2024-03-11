using Microsoft.Office.Interop.Word;
using System.Reflection;
using WordDocCreatorLib.Models;

namespace WordDocCreatorLib
{
    public enum SaveAsDocumentType
    {
        DOC,
        DOCX,
        PDF
    }

    public class WordDocCreator : IDisposable
    {
        private object oMissing = Missing.Value;
        private const string oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        private _Application oWord;
        private _Document oDoc;
        private bool disposedValue;

        public WordDocCreator()
        {
            //Start Word and create a new document.
            oWord = new Application
            {
                Visible = false
            };
            oDoc = oWord.Documents.Add();
        }

        public WordDocCreator(string templateFilePath)
        {
            //Start Word and create a new document.
            oWord = new Application
            {
                Visible = false
            };
            oDoc = oWord.Documents.Add(templateFilePath);
        }
        
        /// <summary>
        /// Updates the text in a region bookmarked in the document with the supplied text. 
        /// Additionally, the text can be hyperlinked if an optional link is provided. 
        /// </summary>
        /// <param name="bookmarkName">The area bookmarked in the document.</param>
        /// <param name="text">The new text to put in the bookmarked region.</param>
        /// <param name="hyperlink">Optional hyperlink to add to the text.</param>
        /// <param name="removeUnderline">Indicates whether to remove the underline that appears under hyperlinked text.</param>
        public void UpdateBookmarkedText(string bookmarkName, string text, string? hyperlink = null, bool removeUnderline = false)
        {
            var bookmarkExists = oDoc.Bookmarks.Exists(bookmarkName);
            if (bookmarkExists)
            {
                var bookmark = oDoc.Bookmarks.get_Item(bookmarkName);
                var range = bookmark.Range;

                if (hyperlink != null)
                {
                    var oHyperlink = range.Hyperlinks.Add(bookmark.Range, hyperlink, TextToDisplay: text);
                    if(removeUnderline) oHyperlink.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                }
                else
                {
                    range.Text = text;
                }
            }
        }

        public void InsertTable(string bookmarkName, WordTable wordTable)
        {
            var bookmarkExists = oDoc.Bookmarks.Exists(bookmarkName);
            if (!bookmarkExists)
            {
                return;
            }

            var bookmark = oDoc.Bookmarks.get_Item(bookmarkName);
            var table = wordTable.Data;
            int rows = table.GetLength(0);
            int columns = table.GetLength(1);

            var oTable = oDoc.Tables.Add(bookmark.Range, rows, columns);
            oTable.Range.Font.Size = 12;
            oTable.Columns.DistributeWidth();
            oTable.set_Style("Table Grid Light");

            for (int i = 0; i < rows; i++)
            {
                for(int j = 0; j < columns; j++)
                {
                    // Office Word Tables indices are not zero based. 
                    oTable.Cell(i + 1, j + 1).Range.Text = table[i,j];
                }
            }

        }

        /// <summary>
        /// Identifies a shape in the document, and simply fills the shape 
        /// with the image supplied at the path. 
        /// </summary>
        /// <param name="alternativeText">The alternative text added to the shape</param>
        /// <param name="imgPath">The path to the image which needs to be inserted into the shape.</param>
        public void FillShapeWithImage(string alternativeText, string imgPath)
        {
            foreach (Shape shape in oDoc.Shapes)
            {
                if (shape.AlternativeText == alternativeText)
                {
                    Console.WriteLine($"Found shape with Alternative Text '{alternativeText}'.");
                    Console.WriteLine("Name: " + shape.Name);
                    Console.WriteLine("ID: " + shape.ID);
                    Console.WriteLine("AnchorID: " + shape.AnchorID);
                    Console.WriteLine("Title: " + shape.Title);
                    Console.WriteLine("Filling shape with supplied image...");
                    shape.Fill.UserPicture(imgPath);
                    Console.WriteLine("Filled shape with supplied image!");
                    break;
                }
            }
        }

        /// <summary>
        /// Saves the document changes at the specified location. 
        /// NOTE: No need to provide the extension. It will be 
        /// automatically added. 
        /// </summary>
        /// <param name="directory">The directory in which the file needs to be saved.</param>
        /// <param name="fileName">The name of the file without extension.</param>
        /// <param name="saveAsDocumentType">The type of file to save as.</param>
        /// <returns>The complete file path (without extension).</returns>
        public string SaveAs(string directory, string fileName, SaveAsDocumentType saveAsDocumentType)
        {
            var wordFileFormat = GetWordFileFormat(saveAsDocumentType);
            var filePath = Path.Combine(directory, fileName);
            oDoc.SaveAs(filePath, wordFileFormat);
            // TODO: I wish to return this path with the extension.
            return filePath;
        }

        private WdSaveFormat GetWordFileFormat(SaveAsDocumentType saveAsDocumentType)
        {
            switch (saveAsDocumentType)
            {
                case SaveAsDocumentType.DOC:
                    return WdSaveFormat.wdFormatDocument97;
                case SaveAsDocumentType.DOCX:
                    return WdSaveFormat.wdFormatDocumentDefault;
                case SaveAsDocumentType.PDF:
                    return WdSaveFormat.wdFormatPDF;
                default:
                    return WdSaveFormat.wdFormatDocumentDefault;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    oDoc?.Close();
                    oWord?.Quit();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        ~WordDocCreator()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
