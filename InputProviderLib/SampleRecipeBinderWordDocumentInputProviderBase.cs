using WordDocCreatorLib.Core.Models;

namespace InputProviderLib
{
    public abstract class SampleRecipeBinderWordDocumentInputProviderBase : IWordDocumentInputProvider
    {
        public abstract IEnumerable<WordDocumentInput> GetWordDocumentInputs();

        public static IEnumerable<WordDocumentInput> GetRecipeBinderSampleInputs()
        {
            return
            [
                GetShengdanyachiBhaajiInput()
            ];
        }

        public static WordDocumentInput GetShengdanyachiBhaajiInput()
        {
            var wordTable = new WordTable(3, 2);
            wordTable.AddData(["Shengdane", "Ardha Kilo"]).AddData(["Oil", "2 tbl. spoons"]).AddData(["Salt", "Chavinusar"]);

            var wordDocumentInput = new WordDocumentInput
            {
                SaveDirectory = "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Automated",
                FileName = "01 Shengdanyachi Bhaaji - Sarika Nilatkar - A5",
                WordTables = new Dictionary<string, WordTable>
                {
                    { "Ingredients_Table", wordTable }
                },
                Images = new Dictionary<string, string>
                {
                    { "Recipe_Image", "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\01 Veg\\01 Shengdanyachi Bhaaji\\IMG_0161.HEIC" }
                },
                Texts = new Dictionary<string, Tuple<string, string>>
                {
                    { "Recipe_Title", Tuple.Create<string, string>("Shengdanyachi Bhaaji", null) },
                    { "Recipe_Author", Tuple.Create("Sarika Nilatkar" , "www.google.com")},
                    { "Procedure", Tuple.Create<string, string>("Shengdane bhajun tyancha kut karun ghyaycha." , null)},
                    { "Recipe_Is_Ready", Tuple.Create<string, string>("Shengdanyachi Bhaaji Tayyar" , null) }
                }
            };

            return wordDocumentInput;
        }

    }
}