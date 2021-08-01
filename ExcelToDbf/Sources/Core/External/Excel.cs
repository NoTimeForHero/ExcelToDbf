using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelToDbf.Sources.Core.External
{
    public class Excel
    {
        Application app;
        public Workbook wb;
        public Worksheet worksheet;
        private List<string> filesToRemove = new List<string>();

        public Excel()
        {
            app = new Application();

        }

        public bool OpenWorksheet(String path)
        {
            // Если не экономим память, то создаём новый экземпляр COM OLE
            wb?.Close(0);

            if (Path.GetExtension(path) == ".csv")
            {
                var convResult = CSV_Converter.Runner.Open(path, ";");
                Logger.info($"Конвертация CSV файла из \"{path}\" в \"{convResult.Filename}\".");
                path = convResult.Filename;
                filesToRemove.Add(convResult.Filename);
                if (!convResult.Success) throw new ApplicationException("Ошибка конвертации файла!");
            }

            wb = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            worksheet = wb.Worksheets[1];
            return true;
        }

        public void close()
        {
            try
            {
                wb?.Close(0);
                app?.Quit();
            }
            finally
            {
                // if (filesToRemove.Count > 0)
                //     Process.Start("explorer.exe", Path.GetDirectoryName(filesToRemove[0]));
                foreach (var file in filesToRemove)
                {
                    File.Delete(file);
                }
            }
        }
    }
}
