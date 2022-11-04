using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using Jint;
using NLog;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelToDbf.Core.Services.Scripts.Context
{
    internal class ExcelContext : AbstractContext
    {
        private readonly ILogger logger;
        private Worksheet worksheet;

        public object GetCellValue(int y, int x)
        {
            // TODO: Чтение данных из кэша

            if (worksheet == null) throw new JSException("Отсутствует лист!");

            try
            {
                return worksheet.Cells[y, x].Value;
            }
            catch (Exception ex)
            {
                logger.Warn($"Ошибка при чтении ячейки x={x},y={y}: {ex.Message}");
                return null;
            }
        }

        public bool Assert(int y, int x, string search)
        {
            return GetCellValue(y,x)?.ToString() == search;
        }

        public ExcelContext(ILogger logger, Engine engine) : base(engine)
        {
            this.logger = logger;
            engine.SetValue("cell", (Func<int, int, object>)GetCellValue);
        }

        public ExcelContext ForDocument(Worksheet worksheet)
        {
            this.worksheet = worksheet;
            return this;
        }

        public object SearchForm(DocForm[] Forms)
        {
            logger.Warn("Ищем форму...");
            return null;
        }
    }
}
