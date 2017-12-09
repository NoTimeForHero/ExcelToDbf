using System;
using System.IO;
using System.Text;

namespace ExcelToDbf.Sources.Core
{

    public class Logger
    {
        #region Variables
        protected StreamWriter writer;
        protected LogLevel level;

        private static readonly Lazy<Logger> lazy = new Lazy<Logger>(() => new Logger());
        public static Logger instance => lazy.Value;
        public static LogLevel Level => instance.level;
        #endregion

        #region Constructor                

        public Logger(string file = null, LogLevel level = LogLevel.INFO)
        {
            this.level = level;
            if (file != null) SetFile(file);
        }
        #endregion

        #region File
        public static void SetFile(string file)
        {
            instance.writer?.Close();
            instance.writer = new StreamWriter(file, false) { AutoFlush = true };
        }
        #endregion

        #region LogLevel       

        public enum LogLevel : byte
        {
            CRITICAL,
            ERROR,
            WARN,
            INFO,
            DEBUG,
            TRACER
        }

        public static void SetLevel(LogLevel newLevel)
        {
            instance.level = newLevel;
        }

        public static void ParseLevel(string newLevel, LogLevel onErrorSet = LogLevel.INFO)
        {
            try
            {
                instance.level = (Logger.LogLevel) Enum.Parse(typeof(Logger.LogLevel), newLevel);
            }
            catch (ArgumentException)
            {
                instance.level = onErrorSet;
            }
        }

        #endregion

        #region Logging Methods

        public static void tracer(object data)
        {
            instance._log(data, LogLevel.TRACER);
        }

        public static void error(object data)
        {
            instance._log(data, LogLevel.ERROR);
        }

        public static void warn(object data)
        {
            instance._log(data, LogLevel.WARN);
        }

        public static void info(object data)
        {
            instance._log(data,LogLevel.INFO);
        }

        public static void debug(object data)
        {
            instance._log(data,LogLevel.DEBUG);
        }

        public static void log(object data, LogLevel level = LogLevel.INFO)
        {
            instance._log(data,level);
        }

        protected void _log(object data, LogLevel curLevel)
        {
            if (curLevel > level) return;

            string prefix = $"[{curLevel}][{DateTime.Now:HH:mm:ss}] ";
            var msg = prefix + data;

            Console.WriteLine(msg);
            if (writer != null)
            {
                msg = msg.Replace("\n", Environment.NewLine + prefix);
                writer.WriteLine(msg);
            }
        }

        #endregion
    }

}
