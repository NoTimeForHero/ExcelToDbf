using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelToDbf.Core.Models;
using Microsoft.Office.Interop.Excel;
using NLog;
using Application = Microsoft.Office.Interop.Excel.Application;
using Point = ExcelToDbf.Core.Models.Point;

namespace ExcelToDbf.Core.Services
{
    public sealed class ExcelService : IDisposable
    {
        private Application app;
        private Workbook workbook;
        private Worksheet worksheet;
        private readonly ConfigProvider config;
        private readonly Dictionary<Point, string> cache = new Dictionary<Point, string>();
        private readonly List<string> temporaryFiles = new List<string>();
        public delegate Cell? HandlerCellGetter(int y, int x);
        public delegate Cell? HandlerRangeFinder(int startY, int endY, int startX, int endX, object expect);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
        public int ProcessId { get; }
        public string Version => app.Version;
        public string SheetName => worksheet.Name;
        public int SheetRows { get; private set; }
        public int SheetColumns { get; private set; }
        public string UsedAddress { get; private set; }
        public string OwnershipDirectory { get; set; }

        public ExcelService(ConfigProvider pvConfig, ILogger logger)
            : this(pvConfig, logger, null) { }

        public ExcelService(ConfigProvider pvConfig, ILogger logger, Action<int> onStarted)
        {
            config = pvConfig;
            var existing = Process.GetProcessesByName("EXCEL").Select(process =>
            {
                using (process) return process.Id;
            }).ToArray();
            app = new Application();
            try
            {
                GetWindowThreadProcessId(new IntPtr(app.Hwnd), out var pid);
                ProcessId = (int)pid;
                if (existing.Contains(ProcessId))
                {
                    Release(app);
                    app = null;
                    throw new InvalidOperationException("Excel не создал отдельный процесс для операции");
                }
                onStarted?.Invoke(ProcessId);
                app.Visible = false;
                app.DisplayAlerts = false;
                app.AskToUpdateLinks = false;
                ((dynamic)app).AutomationSecurity = 3;
            }
            catch { Dispose(); throw; }
        }

        public bool OpenWorksheet(string path)
        {
            CloseWorkbook();
            cache.Clear();
            path = Path.GetFullPath(path);
            if (string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase))
            {
                var converted = CSV_Converter.Runner.Open(path, ";", OwnershipDirectory);
                temporaryFiles.Add(converted.Filename);
                if (!converted.Success) throw new IOException("Ошибка конвертации CSV в Excel");
                path = converted.Filename;
            }
            Workbooks books = null;
            Sheets sheets = null;
            Range used = null;
            Range rows = null;
            Range columns = null;
            try
            {
                books = app.Workbooks;
                workbook = books.Open(path, UpdateLinks: 0, ReadOnly: true, IgnoreReadOnlyRecommended: true,
                    Notify: false, AddToMru: false);
                sheets = workbook.Worksheets;
                worksheet = (Worksheet)sheets[1];
                used = worksheet.UsedRange;
                rows = used.Rows;
                columns = used.Columns;
                SheetRows = used.Row + rows.Count - 1;
                SheetColumns = used.Column + columns.Count - 1;
                UsedAddress = used.Address;
                return true;
            }
            finally { Release(columns); Release(rows); Release(used); Release(sheets); Release(books); }
        }

        public string[] GetSheetNames()
        {
            Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                var names = new string[sheets.Count];
                for (var i = 1; i <= names.Length; i++)
                {
                    Worksheet sheet = null;
                    try { sheet = (Worksheet)sheets[i]; names[i - 1] = sheet.Name; }
                    finally { Release(sheet); }
                }
                return names;
            }
            finally { Release(sheets); }
        }

        // Excel-style one-based matrix, including a single-cell range.
        public object[,] ReadRange(int startY, int endY, int startX, int endX)
        {
            if (worksheet == null) throw new InvalidOperationException("Отсутствует лист");
            if (startY < 1 || startX < 1 || endY < startY || endX < startX || endY > 1048576 || endX > 16384)
                throw new ArgumentOutOfRangeException(nameof(startY), "Недопустимые границы диапазона");
            Range cells = null;
            Range first = null;
            Range last = null;
            Range range = null;
            try
            {
                cells = worksheet.Cells;
                first = (Range)cells[startY, startX];
                last = (Range)cells[endY, endX];
                range = worksheet.Range[first, last];
                object value = range.Value;
                if (value is object[,] matrix) return matrix;
                var single = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
                single[1, 1] = value;
                return single;
            }
            finally { Release(range); Release(last); Release(first); Release(cells); }
        }

        public IEnumerable<object[,]> IterateRanges(int startY, int maxX)
        {
            var size = config.Config.System.BufferSize;
            if (size < 1) throw new ArgumentException("bufferSize должен быть положительным");
            for (var begin = startY; begin <= SheetRows; begin += size)
                yield return ReadRange(begin, Math.Min(begin + size - 1, SheetRows), 1, maxX);
        }

        public Cell? FindRange(int startY, int endY, int startX, int endX, object rawExpect)
        {
            if (!(rawExpect is string expect)) throw new ArgumentException("Ожидаемое значение findRange должно быть строкой");
            var values = ReadRange(startY, endY, startX, endX);
            for (var y = 1; y <= values.GetLength(0); y++)
                for (var x = 1; x <= values.GetLength(1); x++)
                    if (values[y, x]?.ToString() == expect)
                        return new Cell(startY + y - 1, startX + x - 1, expect);
            return null;
        }

        public Cell? GetCellValue(int y, int x)
        {
            var point = new Point(y, x);
            if (!cache.TryGetValue(point, out var value))
            {
                value = ReadRange(y, y, x, x)[1, 1]?.ToString();
                cache[point] = value;
            }
            return new Cell(y, x, value);
        }

        private static void Release(object value)
        {
            // Several Excel getters may return the same RCW (e.g. the selected sheet).
            // Release the reference acquired by this getter, not all aliases at once.
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        private void CloseWorkbook()
        {
            Release(worksheet);
            worksheet = null;
            var closing = workbook;
            workbook = null;
            if (closing == null) return;
            try { closing.Close(false); }
            finally { Release(closing); }
        }

        public void Close() => Dispose();
        public void Dispose()
        {
            try { CloseWorkbook(); }
            finally
            {
                var closing = app;
                app = null;
                try { closing?.Quit(); }
                finally
                {
                    Release(closing);
                    foreach (var path in temporaryFiles) if (File.Exists(path)) File.Delete(path);
                    temporaryFiles.Clear();
                }
            }
        }
    }
}
