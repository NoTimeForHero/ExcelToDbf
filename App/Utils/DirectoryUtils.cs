using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Utils
{
    internal class DirectoryUtils
    {
        public static List<File> GetFilesByExtension(string folder, string[] extensions)
        {
            var results = new List<File>();
            if (string.IsNullOrEmpty(folder)) return results;

            var files = extensions.SelectMany((extension) => Directory.GetFiles(folder, extension, SearchOption.TopDirectoryOnly))
                .Distinct()
                .ToList();

            foreach (var path in files)
            {
                if (path == null) continue;
                var name = Path.GetFileName(path);
                if (name.StartsWith("~$")) continue;
                FileInfo info = new FileInfo(path);
                results.Add(new File
                {
                    FullPath = path,
                    FileName = name,
                    Size = info.Length,
                    Created = info.LastWriteTime
                });
            }
            return results;
        }

        public class File
        {
            public string FullPath { get; set; }
            public string FileName { get; set; }
            public long Size { get; set; }
            public DateTime Created { get; set; }
        }

        public static string BytesToString(long byteCount, string[] suffixes = null)
        {
            suffixes = suffixes ??  new[] { "Б", "Кб", "Мб", "Гб", "Тб" };
            if (byteCount == 0) return "0" + suffixes[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return Math.Sign(byteCount) * num + " " + suffixes[place];
        }
    }
}
