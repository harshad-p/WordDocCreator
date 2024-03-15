namespace WordDocCreatorApp
{
    public interface IWordDocumentInputProvider
    {
        IEnumerable<WordDocumentInput> GetWordDocumentInputs();
    }
}