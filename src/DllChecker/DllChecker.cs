using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;

namespace DllChecker
{
    //https://www.hanselman.com/blog/HowToProgrammaticallyDetectIfAnAssemblyIsCompiledInDebugOrReleaseMode.aspx
    public static class DllChecker
    {
        public static string ScanDirectory(string sDir, StringBuilder sb)
        {
            foreach (string file in Directory.GetFiles(sDir, "*.dll"))
            {
                InspectFile(file, sb);
            }

            var directories = Directory.GetDirectories(sDir);
            if (directories.Length == 0) return sb.ToString();

            foreach (string directory in directories)
            {
                if (directory.EndsWith("roslyn")) continue;
                foreach (string file in Directory.GetFiles(directory, "*.dll"))
                {
                    InspectFile(file, sb);
                }

                ScanDirectory(directory, sb);
            }

            return sb.ToString();
        }

        private static void InspectFile(string file, StringBuilder sb)
        {
            var assembly = Assembly.LoadFrom(file);
            object[] attributes = GetDebuggableAttributes(assembly);
            if (attributes == null) return;
            if (attributes.Length == 0)
            {
                sb.AppendLine($"{file} is RELEASE build");
                return;
            }

            foreach (DebuggableAttribute attr in attributes)
            {
                sb.AppendLine(attr.IsJITOptimizerDisabled
                    ? $"{file} is DEBUG build. Run time optimizer is DISABLED. Run TimeTracking is {attr.IsJITTrackingEnabled}"
                    : $"{file} is RELEASE build. Run time optimizer is ENABLED. Run TimeTracking is {attr.IsJITTrackingEnabled}");
            }
        }

        private static object[] GetDebuggableAttributes(Assembly assembly)
        {
            try
            {
                return assembly.GetCustomAttributes(typeof(DebuggableAttribute), false);
            }
            catch
            {
                return null;
            }
        }
    }
}
