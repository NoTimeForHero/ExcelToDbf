using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelToDbf.Properties;

namespace ExcelToDbf.Sources.Core.Data.FormData
{
    [SuppressMessage("ReSharper", "UnusedMember.Local")]
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Local")]
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    [SuppressMessage("ReSharper", "NotAccessedField.Local")]
    [Serializable]
    public class DataLog
    {
        public enum LogImage : byte
        {
            NONE,
            WARNING,
            ERROR,
            SUCCESS,
            INFO
        }

        protected readonly LogImage type;
        protected readonly string message;

        public String Message => message;
        public Bitmap Image => getImage(type);

        protected Bitmap getImage(LogImage type)
        {
            switch (type)
            {
                case LogImage.NONE:
                    return null;
                case LogImage.WARNING:
                    return Resources.if_warning_16263;
                case LogImage.ERROR:
                    return Resources.if_error_14415;
                case LogImage.SUCCESS:
                    return Resources.if_agt_action_success_3807;
                case LogImage.INFO:
                    return Resources.if_info_3238;
                default:
                    throw new ArgumentException("Unknown image for this LogImage enum: " + type);
            }
        }

        public static List<DataLog> Load()
        {
            if (LastLaunch.Default.LastLog == null) return null;
            byte[] data = Convert.FromBase64String(LastLaunch.Default.LastLog);
            if (data.Length == 0) return null;
            List<DataLog> list;
            using (MemoryStream ms = new MemoryStream(data))
            {
                BinaryFormatter bf = new BinaryFormatter();
                list = bf.Deserialize(ms) as List<DataLog>;
            }
            return list;
        }

        public static void Save(List<DataLog> list)
        {
            byte[] data;
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, list);
                data = ms.ToArray();
            }
            LastLaunch.Default.LastLog = Convert.ToBase64String(data);
            LastLaunch.Default.Save();
        }

        public DataLog(LogImage type, string message)
        {
            this.type = type;
            this.message = message;
        }
    }


    [SuppressMessage("ReSharper", "UnusedMember.Local")]
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Local")]
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    [SuppressMessage("ReSharper", "NotAccessedField.Local")]
    public class DataFileInfo
    {
        protected bool isChecked;
        protected readonly string name;
        protected readonly string size;
        protected readonly string date;
        public readonly string fullPath;

        public delegate void DelegateCheckedChange(bool newState);
        public DelegateCheckedChange CheckedChange;

        public bool Checked
        {
            get => isChecked;
            set
            {
                isChecked = value;
                CheckedChange?.Invoke(value);
            }
        }

        public Bitmap CheckedImg => getImg(isChecked);
        public string Filename => name;
        public string Size => size;
        public string Date => date;

        public DataFileInfo(string fullPath, DelegateCheckedChange CheckedChange = null, string dateFormat = "HH:mm - dd/MM/yyyy")
        {
            this.CheckedChange = CheckedChange;
            this.fullPath = fullPath;

            FileInfo info = new FileInfo(fullPath);
            name = Path.GetFileName(fullPath);
            size = BytesToString(info.Length);
            date = info.LastWriteTime.ToString(dateFormat);
            isChecked = true;
        }

        protected static Bitmap getImg(bool isChecked)
        {
            return isChecked ? Resources.if_checkbox_checked_83249 : Resources.if_checkbox_unchecked_83251;
        }

        protected static String BytesToString(long byteCount)
        {
            string[] suf = { "Б", "Кб", "Мб", "Гб", "Тб" }; //Longs run out around EB
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return Math.Sign(byteCount) * num + " " + suf[place];
        }
    }
}
