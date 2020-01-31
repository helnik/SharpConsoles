using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using Newtonsoft.Json;

namespace ExcelTools
{
    public static class ExcelToJsonConverter
    {
        public static string ConvertToJson(string pathToExcel)
        {
            if (!pathToExcel.HasContext()) throw new ArgumentNullException(nameof(pathToExcel));
            //This connection string works Office 2007+ installed and data is saved in a .xlsx file
            //sets first row as headers HDR=YES
            var connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={pathToExcel};Extended Properties=""Excel 12.0 Xml;HDR=YES""";
            
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                var sheets = GetExcelSheetNames(conn);
                var dic = new Dictionary<string, IEnumerable<object>>();
                foreach (var sheet in sheets)
                {
                    var res = GenerateJson(conn, sheet);
                    if (res.Any()) dic.Add(sheet.TrimEnd(new[] { '$' }), res);
                }

                return JsonConvert.SerializeObject(dic, Formatting.Indented);
            }
        }

        private static List<string> GetExcelSheetNames(OleDbConnection con)
        {
            DataTable dataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            var sheetNames = new List<string>();
            if (dataTable == null) return sheetNames;

            foreach (DataRow row in dataTable.Rows)
            {
                sheetNames.Add(row["TABLE_NAME"].ToString());
            }

            return sheetNames;
        }


        private static List<Dictionary<string, object>> GenerateJson(OleDbConnection conn, string sheetName)
        {
            var cmd = conn.CreateCommand();
            cmd.CommandText = $@"SELECT * FROM [{sheetName}]";


            using (var rdr = cmd.ExecuteReader())
            {
                var query =
                    (from DbDataRecord row in rdr
                     select row).Select(x =>
                     {
                         var item = new Dictionary<string, object>();
                         for (int i = 0; i < x.FieldCount; i++)
                         {
                             item.Add(rdr.GetName(i), x[i]);
                         }
                         return item;
                     });
                return query.ToList();
            }
        }
    }
}
