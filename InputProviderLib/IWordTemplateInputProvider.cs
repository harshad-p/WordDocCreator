using WordDocCreatorLib.Core.Models;

namespace InputProviderLib
{
    public interface IWordTemplateInputProvider
    {
        IEnumerable<WordTemplateInput> GetTemplateInputs();
    }
}