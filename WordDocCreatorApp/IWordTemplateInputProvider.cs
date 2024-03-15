namespace WordDocCreatorApp
{
    public interface IWordTemplateInputProvider
    {
        IEnumerable<WordTemplateInput> GetTemplateInputs();
    }
}