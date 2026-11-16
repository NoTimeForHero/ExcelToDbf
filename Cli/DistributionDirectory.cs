using System;
using System.IO;

namespace ExcelToDbf.Cli
{
    internal static class DistributionDirectory
    {
        public static string Root
        {
            get
            {
                var directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
                if (string.Equals(directory.Name, "DLL", StringComparison.OrdinalIgnoreCase) &&
                    directory.Parent != null && File.Exists(Path.Combine(directory.Parent.FullName, "ExcelToDbf.exe")))
                    return directory.Parent.FullName;
                return directory.FullName;
            }
        }
    }
}
