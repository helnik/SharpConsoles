using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTools
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel's JSON representation:");
            Console.WriteLine();
            string path = ConfigurationManager.AppSettings["outputDirectoryPath"]
                               ?? Directory.GetCurrentDirectory();

            var res = ExcelToJsonConverter.ConvertToJson(path);
            Console.WriteLine(res);
            Console.WriteLine();
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
