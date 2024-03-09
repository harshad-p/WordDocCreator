using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace WordDocCreatorLib
{
    public enum SaveAsDocumentType
    {
        DOC,
        DOCX,
        PDF,
        JPEG
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
        /// Saves the document changes at the specified location. 
        /// NOTE: No need to provide the extension. It will be 
        /// automatically added. 
        /// </summary>
        /// <param name="directory">The directory in which the file needs to be saved.</param>
        /// <param name="fileName">The name of the file without extension.</param>
        /// <param name="saveAsDocumentType">The type of file to save as.</param>
        /// <returns></returns>
        public string SaveAs(string directory, string fileName, SaveAsDocumentType saveAsDocumentType)
        {
            var wordFileFormat = GetWordFileFormat(saveAsDocumentType);
            var filePath = Path.Combine(directory, fileName);
            oDoc.SaveAs(filePath, wordFileFormat);
            
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
