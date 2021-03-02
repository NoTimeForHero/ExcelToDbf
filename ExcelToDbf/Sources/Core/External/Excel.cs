using System;
using System.Collections.Generic;
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
                var csvFilename = CSV_Converter.Runner.Open(@"C:\TEMP\TEST_0203\КО БВ 22.02-27.02.2021 .csv", ";");
                if (csvFilename != null)
                {
                    path = csvFilename;
                    filesToRemove.Add(csvFilename);
                }
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
                foreach (var file in filesToRemove)
                {
                    File.Delete(file);
                }
            }
        }
    }
}
