using WordDocCreatorLib.Core.Models;

namespace InputProviderLib
{
    public interface IWordDocumentInputProvider
    {
        IEnumerable<WordDocumentInput> GetWordDocumentInputs();
    }
}