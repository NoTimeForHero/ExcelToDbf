using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;

namespace ExcelToDbf.Properties {

    [Serializable]
    public sealed class LastLaunch
    {
        private static readonly object syncRoot = new object();
        private static LastLaunch defaultInstance;
        private static XmlSerializer formatter = new XmlSerializer(typeof(LastLaunch));
        public static string Filename = "LastLaunch.xml";

        [XmlIgnore]
        public static LastLaunch Default
        {
            get
            {
                if (defaultInstance == null)
                {
                    lock (syncRoot) if (defaultInstance == null) defaultInstance = Load();
                }
                return defaultInstance;
            }
        }

        public string inputDirectory { get; set; }
        public string outputDirectory { get; set; }
        public string LastLog { get; set; }
        public int positionX { get; set; }
        public int positionY { get; set; }

        public void Save()
        {
            using (var fs = new FileStream(Filename, FileMode.OpenOrCreate))
                formatter.Serialize(fs, this);
        }

        public static LastLaunch Load()
        {
            if (!File.Exists(Filename)) return new LastLaunch();
            try
            {
                using (var fs = new FileStream(Filename, FileMode.OpenOrCreate))
                    return (LastLaunch) formatter.Deserialize(fs);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                return new LastLaunch();
            }
        }

        public LastLaunch() {
            // // Для добавления обработчиков событий для сохранения и изменения параметров раскомментируйте приведенные ниже строки:
            //
            // this.SettingChanging += this.SettingChangingEventHandler;
            //
            // this.SettingsSaving += this.SettingsSavingEventHandler;
            //
        }

    }
}
