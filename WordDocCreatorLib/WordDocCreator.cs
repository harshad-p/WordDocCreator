using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocCreatorLib
{
    public class WordDocCreator
    {
        private readonly string _directory;

        public WordDocCreator(string directory)
        {
            _directory = directory;
        }

        public string SaveAs(string filename)
        {
            var filePath = Path.Combine(_directory, filename);

            return filePath;
        }
    }
}
