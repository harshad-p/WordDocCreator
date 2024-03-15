namespace WordDocCreatorApp
{
    internal class SampleRecipeBinderWordDocumentInputProvider : SampleRecipeBinderWordDocumentInputProviderBase
    {
        public override IEnumerable<WordDocumentInput> GetWordDocumentInputs()
        {
            return new List<WordDocumentInput>(GetRecipeBinderSampleInputs());
        }
    }
}