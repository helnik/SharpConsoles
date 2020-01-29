using System;
using System.Data;
using System.Text;

namespace SimplePOCOGenerator
{
    public static class POCOGenerator
    {
        public static string GenerateClass (Func<IDbConnection> connectionOpener,
            string tableName, string nameSpace)
        {
            ValidateInput(connectionOpener, tableName, nameSpace);  
            
            using (var connection = connectionOpener())
            {   
                var cmd = CreateDbCommand(tableName, connection);

                var reader = cmd.ExecuteReader(CommandBehavior.KeyInfo | CommandBehavior.SingleRow);
                var sb = new StringBuilder(); 

                bool writeClass = true;
                do
                {
                    if (reader.FieldCount <= 1) continue;
                    var schema = reader.GetSchemaTable();
                    foreach (DataRow row in schema.Rows)
                    {
                        if (writeClass)
                        {
                            WriteNameSpaceAndClass(nameSpace, sb, row);
                            writeClass = false;
                        }  
                        WriteProperty(row, sb);
                    }

                    sb.AppendLine("\t}"); 
                    sb.AppendLine("}"); 
                } while (reader.NextResult());

                return sb.ToString(); 
            }
        } 

        private static IDbCommand CreateDbCommand(string tableName, IDbConnection connection)
        {
            var cmd = connection.CreateCommand();
            cmd.CommandText = $"SELECT TOP 1 * FROM {tableName} WITH (NOLOCK)";
            return cmd;
        } 

        private static void ValidateInput(Func<IDbConnection> connectionOpener, string tableName, string nameSpace)
        {
            if (connectionOpener == null) throw new ArgumentNullException(nameof(connectionOpener));
            if (!tableName.HasContext()) throw new ArgumentNullException(nameof(tableName));
            if (!nameSpace.HasContext()) throw new ArgumentNullException(nameof(nameSpace));
        }

        private static void WriteNameSpaceAndClass(string nameSpace, StringBuilder sb, DataRow row)
        {
            sb.AppendLine(@"using System;");
            sb.AppendLine();
            sb.AppendLine($"namespace {nameSpace}");
            sb.AppendLine("{");

            string name = (string)row["BaseTableName"];
            sb.AppendFormat("\tpublic class {0}{1}", name, Environment.NewLine);
            sb.AppendLine("\t{");
        }

        private static void WriteProperty(DataRow row, StringBuilder sb)
        {
            var type = (Type)row["DataType"];
            string typeName = GetTypeName(type);
            bool isNullable = (bool)row["AllowDBNull"] && TypeMapping.NullableTypes.Contains(type);
            string columnName = (string)row["ColumnName"];

            sb.AppendLine(
                $"\t\tpublic {typeName}{(isNullable ? "?" : string.Empty)} {columnName} {{ get; set; }}");
        }

        private static string GetTypeName(Type type)
        {
            TypeMapping.TypeMapToString.TryGetValue(type, out string typeName);
            return typeName.HasContext() ? typeName : type.Name;
        }
    }
}
