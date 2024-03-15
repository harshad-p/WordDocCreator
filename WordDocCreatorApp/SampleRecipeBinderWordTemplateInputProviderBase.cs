namespace WordDocCreatorApp
{
    internal abstract class SampleRecipeBinderWordTemplateInputProviderBase : IWordTemplateInputProvider
    {
        private static readonly IWordDocumentInputProvider _wordDocumentInputProvider;

        static SampleRecipeBinderWordTemplateInputProviderBase()
        {
            _wordDocumentInputProvider = new SampleRecipeBinderWordDocumentInputProvider();
        }

        public abstract IEnumerable<WordTemplateInput> GetTemplateInputs();

        public static IEnumerable<WordTemplateInput> GetRecipeBinderSampleInputs()
        {
            var recipeBinderInput = GetRecipeBinderSampleInput();
            return
            [
                recipeBinderInput
            ];
        }

        public static WordTemplateInput GetRecipeBinderSampleInput()
        {
            var recipeBinderTemplateFilePath = "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Templates\\01 A5 - Kohinoor.dotx";
            return new WordTemplateInput(recipeBinderTemplateFilePath, _wordDocumentInputProvider.GetWordDocumentInputs());
        }

    }
}