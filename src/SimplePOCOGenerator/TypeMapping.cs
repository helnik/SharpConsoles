using System;
using System.Collections.Generic;

namespace SimplePOCOGenerator
{
    internal static class TypeMapping
    {
        internal static readonly Dictionary<Type, string> TypeMapToString = new Dictionary<Type, string>
        {
            
            { typeof(int), "int" },
            { typeof(short), "short" },
            { typeof(byte), "byte" },
            { typeof(byte[]), "byte[]" },
            { typeof(long), "long" },
            { typeof(double), "double" },
            { typeof(decimal), "decimal" },
            { typeof(float), "float" },
            { typeof(bool), "bool" },
            { typeof(string), "string" }
        };

        internal static readonly HashSet<Type> NullableTypes = new HashSet<Type>
        {
            typeof(int),
            typeof(short),
            typeof(long),
            typeof(double),
            typeof(decimal),
            typeof(float),
            typeof(bool),
            typeof(DateTime)
        };
    }
}
