using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NLog;

namespace ExcelToDbf.Utils
{
    internal class FileStorage
    {
        // public static ILogger logger = LogManager.GetCurrentClassLogger();

        public static bool Load<T>(string path, out T target)
        {
            if (!File.Exists(path))
            {
                target = default;
                return false;
            }
            var json = File.ReadAllText(path);
            target = JsonConvert.DeserializeObject<T>(json);
            return true;
        }

        public static void Save<T>(string path, T target)
        {
            var json = JsonConvert.SerializeObject(target, Formatting.Indented);
            File.WriteAllText(path, json);
        }

    }
}
