using WordDocCreatorLib.Models;

namespace WordDocCreatorApp
{
    public class WordTemplateInput
    {
        public WordTemplateInput()
        {
            WordDocumentInputs = new Dictionary<string, IEnumerable<WordDocumentInput>>();
        }

        public IDictionary<string, IEnumerable<WordDocumentInput>> WordDocumentInputs { get; set; }

        public void Add(string templateName, IEnumerable<WordDocumentInput> wordDocumentInput)
        {
            WordDocumentInputs.Add(templateName, wordDocumentInput);
        }

        public static IEnumerable<WordTemplateInput> GetSampleInputs()
        {
            var recipeBinderInput = GetRecipeBinderSampleInput();
            var recipeBinderTemplateInput = new WordTemplateInput();
            recipeBinderTemplateInput.Add(recipeBinderInput.Item1, recipeBinderInput.Item2);

            return new List<WordTemplateInput>
            {
                recipeBinderTemplateInput
            };
        }

        public static Tuple<string, IEnumerable<WordDocumentInput>> GetRecipeBinderSampleInput()
        {
            return Tuple.Create("\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Templates\\01 A5 - Kohinoor.dotx", WordDocumentInput.GetSampleRecipeBinderInputs());
        }

    }

    public class WordDocumentInput
    {
        public string Directory { get; set; }

        public string FileName { get; set; }

        public IDictionary<string, WordTable> WordTables { get; set; }

        public IDictionary<string, string> Images { get; set; }

        public IDictionary<string, Tuple<string, string>> Texts { get; set; }

        public static IEnumerable<WordDocumentInput> GetSampleInputs()
        {
            return new List<WordDocumentInput>(GetSampleRecipeBinderInputs());
        }

        public static IEnumerable<WordDocumentInput> GetSampleRecipeBinderInputs()
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
                Directory = "\\\\Mac\\Home\\Documents\\Recipe Binder\\Workspace\\Automated",
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