using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;


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

            var tableNames = config.GetSection("tableNames").Get<List<string>>();

            foreach (var tableName in tableNames)
            {
                string poco = POCOGenerator.GenerateClass(() =>
                {
                    var c = new SqlConnection(config.GetSection("connectionString")?.Value);
                    c.Open();
                    return c;
                }, tableName, config.GetSection("namespace")?.Value);

                using (var outputFile = new StreamWriter(Path.Combine(directory, $"{tableName}.cs")))
                {
                    outputFile.Write(poco);
                }
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}