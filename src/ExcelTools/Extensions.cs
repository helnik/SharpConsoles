namespace ExcelTools
{
    internal static class Extensions
    {
        internal static bool HasContext(this string sValue) => !string.IsNullOrEmpty(sValue);
    }
}
