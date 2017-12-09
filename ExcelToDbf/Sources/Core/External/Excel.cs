using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelToDbf.Sources.Core.External
{
    public class Excel
    {
        Application app;
        public Workbook wb;
        public Worksheet worksheet;

        public Excel()
        {
            app = new Application();

        }

        public bool OpenWorksheet(String path)
        {
            // Если не экономим память, то создаём новый экземпляр COM OLE
            wb?.Close(0);

            wb = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            worksheet = wb.Worksheets[1];
            return true;
        }

        public void close()
        {
            wb?.Close(0);
            app?.Quit();
        }
    }
}
