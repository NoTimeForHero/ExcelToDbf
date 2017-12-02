using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DomofonExcelToDbf.Sources
{
    class Logger
    {
        bool console = false;
        StreamWriter writer;

        public static Logger instance;

        public Logger(string file = null)
        {
            this.console = (file == null);
            if (file != null)
            {
                writer = new StreamWriter(file, false);
                writer.AutoFlush = true;
            }
        }

        public void log()
        {
            log("");
        }

        public void log(object data)
        {
            Console.WriteLine(data.ToString());
            if (!console)
            {
                writer.WriteLine(data.ToString());
                writer.Flush();
            }
        }

        public void log(string data, object arg0, object arg1 = null, object arg2 = null, object arg3 = null)
        {
            Console.WriteLine(data, arg0, arg1, arg2, arg3);
            if (!console)
            {
                writer.WriteLine(data, arg0, arg1, arg2, arg3);
                writer.Flush();
            }
        }

    }
}
