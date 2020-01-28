using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace SimpleDataContractAttributeGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            string directory = Directory.GetCurrentDirectory();
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(directory)
                .AddJsonFile("appsettings.json", true, true)
                .Build();

            string filesDirectory = config.GetSection("filesToAddDataContractPath")?.Value;
            string outputDirectory = $"{filesDirectory}\\decoratorExport\\";
            if (!Directory.Exists(outputDirectory)) Directory.CreateDirectory(outputDirectory);

            string[] inputFiles = Directory.GetFiles(filesDirectory);
            foreach (string file in inputFiles)
            {
                Console.WriteLine("Starting file(s) decoration...");

                string decorated = SerializationDecorator.DecorateFile(file);
                string fileName = file.Substring(file.LastIndexOf('\\') + 1);
                string outputFile = fileName.EndsWith(".cs") 
                        ? $"{outputDirectory}{fileName}"
                        : $"{outputDirectory}{fileName}.cs";
                WriteFile(decorated, outputFile);

                Console.WriteLine("Completed. Press any key to exit...");
                Console.ReadKey();
            }
        }

        private static void WriteFile(string context, string outputDirectory)
        {

            using (var sw = new StreamWriter(outputDirectory))
            {
                sw.Write(context);
            }
        }
    }
}
