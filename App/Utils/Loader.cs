﻿using System;
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
         public static ILogger logger = LogManager.GetCurrentClassLogger();

        public static bool Load<T>(string path, out T target, T defVal = default)
        {
            try
            {
                return LoadUnsafe(path, out target, defVal);
            }
            catch (Exception ex)
            {
                logger.Warn($"Не удалось загрузить \"{typeof(T).FullName}\" из файла \"${path}\" ");
                logger.Warn(ex.Message);
                target = default;
                return false;
            }
        }

        public static bool LoadUnsafe<T>(string path, out T target, T defVal = default)
        {
            if (!File.Exists(path))
            {
                target = defVal;
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
