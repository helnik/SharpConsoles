using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Dynamic;
using System.IO;
using System.Linq;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;

namespace ExcelTools
{
    public static class ExcelToJson
    {
        public static string ConvertToJson(string pathToExcel)
        {
            //var sheetName = "sheetOne";

            //This connection string works if you have Office 2007+ installed and your 
            //data is saved in a .xlsx file
            var connectionString = String.Format(@"
            Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source={0};
            Extended Properties=""Excel 12.0 Xml;HDR=YES""
        ", pathToExcel);

            //Creating and opening a data connection to the Excel sheet 
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                var sheets = GetExcelSheetNames(conn);

                var sb = new StringBuilder();
                foreach (var sheet in sheets)
                {
                    sb.Append(GenerateJson(conn, sheet));
                }

                return sb.ToString();
            }
        }

        private static string GenerateJson(OleDbConnection conn, string sheetName)
        {
            var cmd = conn.CreateCommand();
            cmd.CommandText = String.Format(
                @"SELECT * FROM [{0}]",//@"SELECT * FROM [{0}$]",
                sheetName
            );


            using (var rdr = cmd.ExecuteReader())
            {
                //dynamic res = new ExpandoObject();
                //Dictionary<string, object> itemDic = new Dictionary<string, object>();
                //for (int i = 0; i <= rdr.FieldCount; i++)
                //{
                //    itemDic.Add(rdr.GetName());
                //}

                //LINQ query - when executed will create anonymous objects for each row
                var query =
                    (from DbDataRecord row in rdr
                        select row).Select(x =>
                    {
                        Dictionary<string, object> item = new Dictionary<string, object>();
                        for (int i = 0; i < x.FieldCount; i++)
                        {
                            //dynamic item = new ExpandoObject();

                            item.Add(rdr.GetName(i), x[i]);
                            
                        }

                        return item;
                    });

                //Generates JSON from the LINQ query
                var json = JsonConvert.SerializeObject(query, Formatting.Indented);
                return json;
            }
        }

        private static List<string> GetExcelSheetNames( OleDbConnection con)
        {
            DataTable dataTable = new DataTable();
            dataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            List<string> sheetNames = new List<string>();
            //for (int i=2; i<dataTable.Rows.Count;)
            foreach (DataRow row in dataTable.Rows)
            {
                if (row["TABLE_NAME"].ToString().Equals("BILLERS$")) continue;
                sheetNames.Add(row["TABLE_NAME"].ToString());
            }

            return sheetNames;
        }

    }

    //class ExcelToJson
    //{
    //    static void Convert(string[] args)
    //    {
    //        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    //        var inFilePath = args[0];
    //        var outFilePath = args[1];

    //        using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
    //        using (var outFile = File.CreateText(outFilePath))
    //        {
    //            using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
    //            { FallbackEncoding = Encoding.GetEncoding(1252) }))
    //            using (var writer = new JsonTextWriter(outFile))
    //            {
    //                writer.Formatting = Formatting.Indented; //I likes it tidy
    //                writer.WriteStartArray();
    //                reader.Read(); //SKIP FIRST ROW, it's TITLES.
    //                do
    //                {
    //                    while (reader.Read())
    //                    {
    //                        //peek ahead? Bail before we start anything so we don't get an empty object
    //                        var status = reader.GetString(0);
    //                        if (string.IsNullOrEmpty(status)) break;

    //                        writer.WriteStartObject();
    //                        writer.WritePropertyName("Status");
    //                        writer.WriteValue(status);

    //                        writer.WritePropertyName("Title");
    //                        writer.WriteValue(reader.GetString(1));

    //                        writer.WritePropertyName("Host");
    //                        writer.WriteValue(reader.GetString(6));

    //                        writer.WritePropertyName("Guest");
    //                        writer.WriteValue(reader.GetString(7));

    //                        writer.WritePropertyName("Live");
    //                        writer.WriteValue(reader.GetDateTime(5));

    //                        writer.WritePropertyName("Url");
    //                        writer.WriteValue(reader.GetString(11));

    //                        writer.WritePropertyName("EmbedUrl");
    //                        writer.WriteValue($"{reader.GetString(11)}player");
    //                        /*
    //                        <iframe src="https://channel9.msdn.com/Shows/Azure-Friday/Erich-Gamma-introduces-us-to-Visual-Studio-Online-integrated-with-the-Windows-Azure-Portal-Part-1/player" width="960" height="540" allowFullScreen frameBorder="0"></iframe>
    //                         */

    //                        writer.WriteEndObject();
    //                    }
    //                } while (reader.NextResult());
    //                writer.WriteEndArray();
    //            }
    //        }
    //    }
    //}
}
