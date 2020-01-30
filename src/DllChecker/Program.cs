using System.Configuration;
using System.IO;
using System.Text;

namespace DllChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            string directory = ConfigurationManager.AppSettings["directoryPath"]
                               ?? Directory.GetCurrentDirectory();
            var sb = new StringBuilder();
            string result = DllChecker.ScanDirectory(directory, sb);

            using (var outputFile = new StreamWriter(
                Path.Combine(Directory.GetCurrentDirectory(), "result.txt")))
            {
                outputFile.Write(result);
            }
        }
    }
}

