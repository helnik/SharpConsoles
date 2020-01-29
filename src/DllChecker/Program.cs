using System.Configuration;
using System.IO;

namespace DllChecker
{
    class Program
    {
        
        static void Main(string[] args)
        {
            string directory = ConfigurationManager.AppSettings["directoryPath"] 
                               ?? Directory.GetCurrentDirectory();
            string result = DllChecker.ScanDirectory(directory);

            using (var outputFile = new StreamWriter(
                Path.Combine(Directory.GetCurrentDirectory(), "result.txt")))
            {
                outputFile.Write(result);
            }
        }
    }
}

