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
using Point = ExcelToDbf.Core.Models.Point;

namespace ExcelToDbf.Core.Services
{
    public class ExcelService : IDisposable
    {
        private readonly Application app;
        private readonly List<string> filesToRemove = new List<string>();
        private readonly Dictionary<Point, string> cacheCells = new Dictionary<Point, string>();
        private readonly ConfigProvider pvConfig;
        private readonly ILogger logger;

        public Workbook wb { get; private set; }
        public Worksheet worksheet { get; private set; }

        public delegate Cell? HandlerCellGetter(int y, int x);
        public delegate Cell? HandlerRangeFinder(int startY, int endY, int startX, int endX, object expect);

        public int SheetRows => worksheet.UsedRange.Rows.Count;

        public ExcelService(ConfigProvider pvConfig, ILogger logger)
        {
            app = new Application();
            if (Constants.ExcelDebug) app.Visible = true;
            this.logger = logger;
            this.pvConfig = pvConfig;
        }

        public bool OpenWorksheet(string path)
        {
            wb?.Close(0);
            cacheCells.Clear();

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

        public IEnumerable<object[,]> IterateRanges(int startY, int maxX)
        {
            var bufferSize = pvConfig.Config.System.BufferSize;
            int begin = startY - 1;
            int end = begin + bufferSize;
            bool EOF = false;

            int maxY = SheetRows;
            logger.Info($"Размер документа {maxY} строк!");

            while (!EOF)
            {
                if (end > maxY) end = maxY;
                var range_start = worksheet.Cells[begin + 1, 1];
                var range_end = worksheet.Cells[end + 1, maxX];
                logger.Trace($"Загрузка блока {begin}-{end}");
                var range = worksheet.Range[range_start, range_end];
                var value = range.Value;
                yield return value;
                begin += bufferSize + 1;
                end += bufferSize + 1;
                if (begin > maxY)
                {
                    logger.Debug($"Выход по границе документа: {begin} > {maxY}");
                    EOF = true;
                }
            }
        }

        public Cell? FindRange(int startY, int endY, int startX, int endX, object rawExpect)
        {
            if (worksheet == null) throw new InvalidOperationException("Отсутствует лист!");

            if (!(rawExpect is string expect)) throw new ArgumentException("Аргумент FindRange должен быть строкой!");

            var range_start = worksheet.Cells[startY, startX];
            var range_end = worksheet.Cells[endY, endX];
            object[,] range = worksheet.Range[range_start, range_end].Value;

            var lenY = endY - startY;
            var lenX = endX - startX;

            for (int y = 1; y <= lenY + 1; y++)
            {
                for (int x = 1; x <= lenX + 1; x++)
                {
                    var got = range[y, x]?.ToString();
                    if (expect == got)
                    {
                        var cell = new Cell(startY + y - 1, startX + x - 1, got);
                        return cell;
                    }
                }
            }

            return null;
        }

        public Cell? GetCellValue(int y, int x)
        {
            if (worksheet == null) throw new InvalidOperationException("Отсутствует лист!");

            var point = new Point(y, x);
            if (cacheCells.TryGetValue(point, out var cacheValue)) return new Cell(y, x, cacheValue);

            try
            {
                var value = worksheet.Cells[y, x].Value;
                cacheCells[point] = value;
                return new Cell(y, x, value);
            }
            catch (Exception ex)
            {
                logger.Warn($"Ошибка при чтении ячейки x={x},y={y}: {ex.Message}");
                return new Cell(y, x);
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
