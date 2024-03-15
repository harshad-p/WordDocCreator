using WordDocCreatorLib.Core.Models;

namespace InputProviderLib
{
    public class SampleRecipeBinderWordTemplateInputProvider : SampleRecipeBinderWordTemplateInputProviderBase
    {
        public override IEnumerable<WordTemplateInput> GetTemplateInputs()
        {
            return new List<WordTemplateInput>(GetRecipeBinderSampleInputs());
        }
    }
}