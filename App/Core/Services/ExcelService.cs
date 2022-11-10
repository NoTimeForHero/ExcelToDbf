using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils.Extensions;
using Microsoft.Office.Interop.Excel;
using NLog;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelToDbf.Core.Services
{
    public class ExcelService : IDisposable
    {
        private readonly Application app;
        private readonly List<string> filesToRemove = new List<string>();
        private readonly ILogger logger;

        public Workbook wb { get; private set; }
        public Worksheet worksheet { get; private set; }

        public delegate Cell? HandlerCellGetter(int y, int x);

        public ExcelService(ILogger logger)
        {
            app = new Application();
            if (Constants.ExcelDebug) app.Visible = true;
            this.logger = logger;
        }

        public bool OpenWorksheet(string path)
        {
            wb?.Close(0);

            if (Path.GetExtension(path) == ".csv")
            {
                var convResult = CSV_Converter.Runner.Open(path, ";");
                logger.Info($"Конвертация CSV файла из \"{path}\" в \"{convResult.Filename}\".");
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

        public Cell? GetCellValue(int y, int x)
        {
            // TODO: Чтение данных из кэша

            if (worksheet == null) throw new InvalidOperationException("Отсутствует лист!");

            try
            {
                return new Cell
                {
                    Y = y,
                    X = x,
                    Value = worksheet.Cells[y, x].Value
                };
            }
            catch (Exception ex)
            {
                logger.Warn($"Ошибка при чтении ячейки x={x},y={y}: {ex.Message}");
                return new Cell { Y = y, X = x }; ;
            }
        }

        public void Close()
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
                logger.Debug("Удаление временных файлов: " + filesToRemove.JoinString(", "));
                foreach (var file in filesToRemove)
                {
                    File.Delete(file);
                }
            }
        }

        public void Dispose()
        {
            Close();
        }
    }
}
