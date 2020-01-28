using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace SimplePOCOGenerator
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

            string tableName = config.GetSection("tableName")?.Value;
            string outputClassName = config.GetSection("outputClassName")?.Value;

            string poco = POCOGenerator.GenerateClass(() =>
            {
                var c = new SqlConnection(config.GetSection("connectionString")?.Value);
                c.Open();
                return c;
            }, tableName, config.GetSection("namespace")?.Value, outputClassName);

            string fileName = outputClassName.HasContext() ? outputClassName : tableName;
            using (var outputFile = new StreamWriter(Path.Combine(directory, $"{fileName}.cs")))
            {
                outputFile.Write(poco); 
            }

            Console.WriteLine(poco);
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
