using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace WordDocCreatorLib
{
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

        public string SaveAs(string filePath)
        {
            var filePathLocal = Path.Combine(filePath + ".docx");
            oDoc.SaveAs(filePathLocal);
            return filePath;
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
