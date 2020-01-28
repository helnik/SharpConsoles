using System.Collections.Generic;
using System.IO;

namespace SimpleDataContractAttributeGenerator
{
    public static class SerializationDecorator
    {
        public static string DecorateFile(string file)
        {
            string fileContext = ReadFile(file);
            List<string> splitContext = fileContext.SplitWithString("public");

            string usings = "using System.Runtime.Serialization;\r\n" + splitContext[0].SplitWithString("namespace")[0];
            string nameSpace = $"namespace {splitContext[0].SplitWithString("namespace")[1]}";
            splitContext.RemoveAt(0); // remove usings
            string decoratedContext = Decorate(splitContext);
            return usings + nameSpace + decoratedContext;
        }

        private static string ReadFile(string file)
        {
            using (var sr = new StreamReader(file))
            {
                return sr.ReadToEnd();
            }
        }

        private static string Decorate(List<string> properties)
        {
            string context = string.Empty;
            foreach (var property in properties)
            {
                if (property.Contains("class"))
                {
                    context += $"\r\n    [DataContract]\r    public{property}";
                }
                else
                {
                    string propertyName = property.TrimStart(' ').Split(" ")[1].ToLowerFirstChar();
                    context += $"[DataMember(Name = \"{propertyName}\")]\r        public{property}";
                }
            }
            return context;
        }
    }
}
