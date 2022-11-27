using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts.Data;
using ExcelToDbf.Utils.Extensions;
using Jint;
using Jint.Native;
using Jint.Native.Array;
using Jint.Native.Function;
using Jint.Native.Object;
using NLog;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using JintSerializer = Jint.Native.Json.JsonSerializer;

namespace ExcelToDbf.Core.Services.Scripts.Context
{
    public partial class ExcelContext : AbstractContext
    {
        private readonly ILogger logger;
        private readonly IConfigContext config;
        private ExcelService.HandlerCellGetter cellValueGetter = (y, x) => throw new ArgumentNullException(nameof(cellValueGetter));
        private ExcelService.HandlerRangeFinder rangeFinderGetter = (y1, y2, x1, x2, exp) => throw new ArgumentNullException(nameof(rangeFinderGetter));
        private readonly JintSerializer parser;
        private ObjectInstance currentContext;

        public ExcelContext(ILogger logger, IConfigContext config, Engine engine) : base(engine)
        {
            this.logger = logger;
            this.config = config;
            parser = new JintSerializer(engine);
        }

        public ExcelContext Connect(ExcelService.HandlerCellGetter getter, ExcelService.HandlerRangeFinder getter2)
        {
            cellValueGetter = getter;
            rangeFinderGetter = getter2;
            currentContext = new ObjectInstance(engine);
            engine.SetValue("context", currentContext);
            return this;
        }

        public bool TryGetContextValue<T>(string keyName, out T result)
        {
            var value = currentContext.Get(keyName);
            if (value is JsUndefined)
            {
                result = default;
                return false;
            }
            result = parser.Deserialize<T>(value);
            return true;
        }

        public SearchFormResult SearchForm(FileModel file)
        {
            var result = new SearchFormResult();
            foreach (var form in config.Forms)
            {
                logger.Info($"Проверяем форму: {form.Name}");

                var matches = new List<SearchMatch>();
                result.Report[form] = matches;

                engine.SetValue("cell", cellValueGetter);
                Func<object, object, bool> ContextAssert = (got, expect) => Assert(matches, got, expect);
                engine.SetValue("assert", ContextAssert);
                engine.SetValue("findRange", rangeFinderGetter);

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
                    if (config.Data.Config.System.FastSearch) return result;
                }
            }
            return result;
        }

        private bool Assert(List<SearchMatch> matches, object got, object rawExpect)
        {
            string expect = rawExpect?.ToString();
            Predicate<string> comparer;
            switch (rawExpect)
            {
                case string expString:
                    comparer = x => expString == x;
                    break;
                case Regex expRegex:
                    comparer = x => expRegex.IsMatch(x);
                    break;
                default:
                    throw new InvalidOperationException($"Unknown assert 'Expect' type: {rawExpect?.GetType().FullName}");
            }
            SearchMatch match;
            switch (got)
            {
                case string simple:
                    match = SearchMatch.Make(expect, simple, comparer(simple));
                    break;
                case Cell cell:
                    match = SearchMatch.Make(expect, cell.Value, comparer(cell.Value)).With(cell.Y, cell.X);
                    break;
                default:
                    throw new InvalidOperationException($"Unknown assert 'Got' type: {got.GetType().FullName}");
            }
            logger.Trace(match.ToString());
            matches.Add(match);
            if (config.Data.Config.System.FastSearch && !match.Matches) throw new StopFunctionException();
            return match.Matches;
        }

        private void RunStop()
        {
            throw new StopFunctionException();
        }

        public Dictionary<string, object> Transform(DocForm form, object[] record)
        {
            engine.SetValue("stop", (Action)RunStop);
            var value = engine.Invoke(form.Write, new object[]{ record });
            return parser.Deserialize<Dictionary<string,object>>(value);
        }

        public void CallHook(DocForm form, DocForm.HookType type)
        {
            ScriptFunctionInstance target;
            switch (type)
            {
                case DocForm.HookType.Before:
                    target = form.BeforeWrite;
                    break;
                case DocForm.HookType.After:
                    target = form.AfterWrite;
                    break;
                default:
                    throw new NotImplementedException(type.ToString());
            }
            if (target == null)
            {
                logger.Debug($"Отсутствует хук {type} для формы \"{form.Name}\"!");
                return;
            }
            engine.Invoke(target);
        }
    }
}
