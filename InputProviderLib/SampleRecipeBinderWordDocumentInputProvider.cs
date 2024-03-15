using WordDocCreatorLib.Core.Models;

namespace InputProviderLib
{
    public class SampleRecipeBinderWordDocumentInputProvider : SampleRecipeBinderWordDocumentInputProviderBase
    {
        public override IEnumerable<WordDocumentInput> GetWordDocumentInputs()
        {
            return new List<WordDocumentInput>(GetRecipeBinderSampleInputs());
        }
    }
}