using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelTools
{
    class Program
    {
        static void Main(string[] args)
        {
            string outputDirectoryPath = ConfigurationManager.AppSettings["outputDirectoryPath"]
                               ?? Directory.GetCurrentDirectory();
            if (!outputDirectoryPath.EndsWith(@"\"))
                outputDirectoryPath += @"\";

            CreateDummyExcel(outputDirectoryPath);

            var result = ExcelToJson.ConvertToJson($@"{outputDirectoryPath}\test.xlsx");
            Console.WriteLine(result);
            Console.ReadKey();
        }

        

        private static void CreateDummyExcel(string outputDirectoryPath)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                CreateBillersList(excel);

                AddBiller(excel, "ΔΕΗ");

                AddHyperLinks(excel);

                FileInfo excelFile = new FileInfo($@"{outputDirectoryPath}\test.xlsx");
                excel.SaveAs(excelFile);

                var result = ExcelToJson.ConvertToJson($@"{outputDirectoryPath}\testFill.xlsx");
                Console.WriteLine(result);
                Console.ReadKey();
            }
        }

        private static void AddBiller(ExcelPackage excel, string worksheetName)
        {
            excel.Workbook.Worksheets.Add(worksheetName);
            var excelWorksheet = excel.Workbook.Worksheets[worksheetName];
            List<string[]> headerRow = new List<string[]>()
            {
                new string[] {"ID", "First Name", "Last Name"}
            };
            string headerRange = "A1:" + char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
            excelWorksheet.Cells[headerRange].LoadFromArrays(headerRow);
            excelWorksheet.View.FreezePanes(2, headerRow[0].Length); //bazw to prwto unfrozen row
        }

        private static void CreateBillersList(ExcelPackage excel)
        {
            excel.Workbook.Worksheets.Add("BILLERS");
            var excelWorksheet = excel.Workbook.Worksheets["BILLERS"];
            excelWorksheet.Cells["A1"].LoadFromText("AVAILABLE BILLERS");
            var cellData = new List<object[]>
            {
                new object[] {"ΔΕΗ"},
                new object[] {"ΕΥΔΑΠ"},
                new object[] {"ΟΤΕ"}
            };
            excelWorksheet.Cells[2, 1].LoadFromArrays(cellData);
            excelWorksheet.View.FreezePanes(2, 1); //bazw to prwto un frozen row
        }

        private static void AddHyperLinks(ExcelPackage excel)
        {
            var billersList = excel.Workbook.Worksheets["BILLERS"];
            //thelw na exw to count me tous billers gia na kanw iteration
            int rowCount = 2;//< your specific row>
            Uri url = new Uri("#'ΔΕΗ'!A1", UriKind.Relative);
            billersList.Cells[rowCount, 1].Hyperlink = url;
                //new ExcelHyperLink((char)39 +
                //                   "ΔΕΗ" +
                //                   (char)39 + 
                //                   "!A1(specific cell on that sheet)", billersList.Cells[rowCount, 1].Text);
            
        }

        private static void AddHyperLink(ExcelWorksheet ws, ExcelRange source, ExcelRange destination, string displayText)
        {
            source.Formula = "HYPERLINK(\"[\"&MID(CELL(\"filename\"),SEARCH(\"[\",CELL(\"filename\"))+1, SEARCH(\"]\",CELL(\"filename\"))-SEARCH(\"[\",CELL(\"filename\"))-1)&\"]" + destination.FullAddress + "\",\"" + displayText + "\")";
        }
    }
}
