﻿namespace WordDocCreatorApp
{
    internal class SampleRecipeBinderWordTemplateInputProvider : SampleRecipeBinderWordTemplateInputProviderBase
    {
        public override IEnumerable<WordTemplateInput> GetTemplateInputs()
        {
            return new List<WordTemplateInput>(GetRecipeBinderSampleInputs());
        }
    }
}