using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DomofonExcelToDbf.Sources
{
    class Excel
    {
        Application app;
        Workbook wb;
        public Worksheet worksheet;
        protected bool saveMemory;

        public Excel(bool saveMemory)
        {
            if (saveMemory) app = new Application();
            this.saveMemory = saveMemory;

        }

        public bool OpenWorksheet(String path)
        {
            // Если не экономим память, то создаём новый экземпляр COM OLE
            if (saveMemory)
            {
                wb?.Close(0);
            }
            else
            {
                app?.Quit();
                app = new Application();
            }

            wb = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            if (wb.Worksheets.Count < 1)
            {
                Logger.instance.log("Выбранный Excel не содержит ни одного листа!");
                return false;
            }

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
