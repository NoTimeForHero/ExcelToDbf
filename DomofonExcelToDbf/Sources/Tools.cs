using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DomofonExcelToDbf.Sources.Xml;

namespace DomofonExcelToDbf.Sources
{
    class Tools
    {
        // <summary>
        // Метод считывает внутренний ресурс и записывает его в файл, возвращая статус существования ресурса
        // </summary>
        // <param name="resourceName">Имя внутренного ресурса</param>
        // <param name="fileName">Имя внутренного ресурса</param>
        // <returns>false если внутренний ресурс не был найден</returns>
        public static bool WriteResourceToFile(string resourceName, string fileName)
        {
            using (var resource = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (resource == null) return false;
                using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    resource.CopyTo(file);
                }
            }
            return true;
        }

    }

    public class MyException : Exception
    {
        private string myStackTrace;

        public MyException(string message, Exception exp) : base(message)
        {
            this.myStackTrace = exp.StackTrace;
        }

        public override string StackTrace
        {
            get
            {
                return base.StackTrace + "\n" + myStackTrace;
            }
        }
    }
}
