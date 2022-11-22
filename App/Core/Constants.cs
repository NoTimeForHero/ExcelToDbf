using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Core
{
    internal static class Constants
    {
        public static readonly string ApplicationTitle = "Конвертирование Excel документов в DBF";
        public static readonly string SettingsFile = "config.js";
        public static readonly string PreloadFile = "Preload.json";
        public static readonly bool ExcelDebug = false;
        public static readonly string LastLaunchFile = "LastLaunch.json";
    }
}
