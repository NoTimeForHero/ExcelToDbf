using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    public class ExcelContext : AbstractContext
    {
        private readonly ILogger logger;
        private readonly IConfigContext config;
        private ExcelService.HandlerCellGetter cellValueGetter = (y, x) => throw new ArgumentNullException(nameof(cellValueGetter));

        public ExcelContext(ILogger logger, IConfigContext config, Engine engine) : base(engine)
        {
            this.logger = logger;
            this.config = config;
        }

        public ExcelContext Connect(ExcelService.HandlerCellGetter getter)
        {
            cellValueGetter = getter;
            return this;
        }

        public SearchFormResult SearchForm(FileModel file)
        {
            var result = new SearchFormResult { Target = file };
            foreach (var form in config.Forms)
            {
                logger.Info($"Проверяем форму: {form.Name}");

                var matches = new List<SearchMatch>();
                result.Report[form] = matches;

                Action<object, string> ContextAssert = (got, expect) =>
                {
                    SearchMatch match;
                    switch (got)
                    {
                        case string simple:
                            match = SearchMatch.Make(expect, simple, expect == simple);
                            break;
                        case Cell cell:
                            match = SearchMatch.Make(expect, cell.Value, expect == cell.Value).With(cell.Y, cell.X);
                            break;
                        default:
                            throw new InvalidOperationException($"Unknown assert value type: {got.GetType().FullName}");
                    }
                    logger.Trace(match.ToString());
                    matches.Add(match);
                    if (config.Data.System.FastSearch && !match.Matches) throw new StopFunctionException();
                };
                engine.SetValue("cell", cellValueGetter);
                engine.SetValue("assert", ContextAssert);

                try
                {
                    form.Rules.Call();
                }
                catch (StopFunctionException)
                {
                    logger.Info($"Форма \"{form.Name}\" не подходит по условию!");
                }

                var isMatches = matches.All(x => x.Matches);
                if (isMatches)
                {
                    logger.Info($"Форма \"{form.Name}\" подходит для документа \"{file.FileName}\"!");
                    if (result.Result == null) result.Result = form;
                    if (config.Data.System.FastSearch) return result;
                }
            }
            return result;
        }

        private class StopFunctionException : Exception
        {

        }
    }
}
