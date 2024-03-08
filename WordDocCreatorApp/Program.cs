using WordDocCreatorLib;

namespace WordDocCreatorApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            var directory = "./";
            var wordDocCreator = new WordDocCreator(directory);
            var fileName = "Test-Document";
            Console.WriteLine(wordDocCreator.SaveAs(fileName));
        }
    }
}
