﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelToDbf.Sources.Core.Data.FormData;
using ExcelToDbf.Sources.Core.Data.TData;
using ExcelToDbf.Sources.Core.Data.Xml;
using Jint.Native;
using Microsoft.Office.Interop.Excel;

namespace ExcelToDbf.Sources.Core
{
    public class Work
    {
        protected Dictionary<string, TVariable> staticVars = new Dictionary<string, TVariable>();
        protected Dictionary<string, TVariable> dynamicVars = new Dictionary<string, TVariable>();
        protected HashSet<TCondition> conditions = new HashSet<TCondition>();

        protected int buffer;
        protected int startY;
        protected int endX;
        protected List<Xml_Validator> validators;
        protected Worksheet worksheet;
        protected string jScript;

        public int StartY => startY;

        // Переменные для нахождения номера строки и переменной исключения
        protected int total;

        protected TVariable exception_var;

        public Dictionary<string, TVariable> stepScope = new Dictionary<string, TVariable>();
        public Action<DataLog.LogImage, string> guiLogger;

        public Work(Worksheet worksheet, Xml_Form form, int buffer)
        {
            this.worksheet = worksheet;
            this.buffer = buffer;

            InitVariables(form);
            startY = findStartY(form);
            endX = form.Fields.EndX;
            validators = form.Validate;
            jScript = form.Fields.Script;
        }

        public TimeSpan IterateRecords(Action<Dictionary<string, TVariable>> callback, Action<int> guiCallback = null)
        {
            if (buffer <= 0) throw new ArgumentException("Буфер обработки должен быть больше ноля!");
            total = 0;
            Stopwatch watch = Stopwatch.StartNew();
            try
            {
                __IterateRecords(worksheet, callback, guiCallback);
            }
            catch (Exception ex) when (!Program.DEBUG)
            {
                if (exception_var == null) throw;
                string message = $"Ошибка на строке {startY + total - 1}, ячейке {exception_var.x} в переменной {exception_var.name}:\n{ex.Message}";
                throw new MyException(message, ex);
            }
            watch.Stop();
            FinalChecks();

            return watch.Elapsed;
        }

        protected int findStartY(Xml_Form form)
        {
            var target = form.Fields.StartY;
            if (target.IsSimple && target.SimpleValue.HasValue)
            {
                return target.SimpleValue.Value;
            }

            var group = form.Rules.OfType<Xml_Equal_Group>().FirstOrDefault(x => x.Name == target.group.name);
            if (group == null) throw new InvalidOperationException($"Не удалось найти группу с именем '{target.group.name}'");

            var point = Program.XmlTools.findGroupPosition(worksheet, group);
            if (point == null) throw new InvalidOperationException($"Не удалось найти координаты для группы '{group.Name}'!");

            if (target.group.position != "after")
            {
                throw new ArgumentException($"Недопустимое значение 'position': {target.group.position}",
                    nameof(target.group.position));
            }
            return point.Value.Y + target.group.Y;
        }

        [SuppressMessage("ReSharper", "ConditionIsAlwaysTrueOrFalse")]
        protected void __IterateRecords(Worksheet worksheet, Action<Dictionary<string, TVariable>> callback, Action<int> guiCallback = null)
        {
            int begin = startY;
            int end = startY + buffer;

            var maxY = worksheet.UsedRange.Rows.Count;

            bool EOF = false;

            Stopwatch watch = Stopwatch.StartNew();
            stepScope.Clear();
            foreach (var var in staticVars.Values)
                SetVar(var, worksheet.Cells[var.y, var.x].Value);
            watch.Stop();
            Logger.debug("Заполнение массива локальных переменных: " + watch.ElapsedMilliseconds);

            int i = 0;
            bool skipRecord = false;
            bool stopLoop = false;
            object[,] tmp = null;

            var engine = new Jint.Engine();
            engine.SetValue("skipRecord", (System.Action)(() => skipRecord = true));
            engine.SetValue("stopLoop", (System.Action)(() => stopLoop = true));
            engine.SetValue("throwError", (Action<string>) ((message) =>
                throw new ApplicationException($"JS ошибка: {message}")));
            engine.SetValue("log", (Action<object>)Logger.info);
            engine.SetValue("info", (Action<string>) ((msg) => guiLogger?.Invoke(DataLog.LogImage.INFO, msg)));
            engine.SetValue("warn", (Action<string>)((msg) => guiLogger?.Invoke(DataLog.LogImage.WARNING, msg)));
            engine.SetValue("getRow", (Func<object>)(() => tmp?.GetRow(i, 1)));
            var getCache = new Dictionary<string, object>();
            engine.SetValue("cell", (Func<int,int,object>)((y, x) =>
            {
                bool hasValue = getCache.TryGetValue($"{y}/{x}", out var obj);
                if (hasValue) return obj;
                obj = worksheet.Cells[y, x]?.Value;
                getCache[$"{y}/{x}"] = obj;
                return obj;
            }));
            engine.SetValue("set", (Func<string,object,bool>)((name,value) =>
            {
                TVariable tVar = null;
                bool hasValue = false;
                if (!hasValue) hasValue = dynamicVars.TryGetValue(name, out tVar);
                if (!hasValue) hasValue = staticVars.TryGetValue(name, out tVar);
                if (hasValue) SetVar(tVar, value);
                return hasValue;
            }));
            engine.SetValue("get", (Func<string, object>)((name) =>
            {
                TVariable tVar = null;
                bool hasValue = false;
                if (!hasValue) hasValue = dynamicVars.TryGetValue(name, out tVar);
                if (!hasValue) hasValue = staticVars.TryGetValue(name, out tVar);
                if (hasValue) return tVar.value;
                return null;
            }));

            JsValue scriptEveryRow = null;
            JsValue scriptBefore = null;
            JsValue scriptAfter = null;

            if (jScript != null)
            {
                var jVal = engine.Execute(JSHelper.decodeXMLEntities(jScript)).GetCompletionValue();
                if (jVal.IsObject())
                {
                    var jObj = jVal.AsObject();
                    scriptEveryRow = jObj.getOrDefault("row", null);
                    scriptAfter = jObj.getOrDefault("after", null);
                    scriptBefore = jObj.getOrDefault("before", null);
                }
            }

            Stopwatch watchTotal = Stopwatch.StartNew();
            scriptBefore?.Invoke();
            while (!EOF)
            {
                var range_start = worksheet.Cells[begin, 1];
                var range_end = worksheet.Cells[end, endX];
                var range = worksheet.Range[range_start, range_end];
                tmp = range.Value;

                watch = Stopwatch.StartNew();
                for (i = 1; i <= buffer; i++)
                {
                    skipRecord = false;
                    stopLoop = false;
                    total++;
                    foreach (TCondition cond in conditions)
                    {
                        var cellValue = tmp[i, cond.x]?.ToString() ?? "";
                        bool equal = false;
                        if (cond.isRegex) equal = Regex.Match(cellValue, cond.mustBe).Success;
                        else equal = cellValue == cond.mustBe;
                        if (equal)
                        {
                            foreach (TAction item in cond.onTrue)
                            {
                                switch (item)
                                {
                                    case TInterrupt tinter:
                                        switch (tinter.action)
                                        {
                                            case TInterrupt.Action.SKIP_RECORD:
                                                Logger.tracer($"Пропуск записи по условию: значение в ячейке x={cond.x} равно {cond.mustBe}");
                                                skipRecord = true;
                                                break;
                                            case TInterrupt.Action.STOP_LOOP:
                                                Logger.tracer($"Выход из цикла по условию: значение в ячейке x={cond.x} равно {cond.mustBe}");
                                                stopLoop = true;
                                                break;
                                        }
                                        continue;
                                    case TVariable var:
                                        SetVar(var, tmp[i, var.x]);
                                        break;
                                }
                            }
                        }
                        else
                        {
                            foreach (TAction item in cond.onFalse)
                            {
                                switch (item)
                                {
                                    case TInterrupt tinter:
                                        switch (tinter.action)
                                        {
                                            case TInterrupt.Action.SKIP_RECORD:
                                                Logger.tracer($"Пропуск записи по условию: значение в ячейке x={cond.x} равно {cond.mustBe}");
                                                skipRecord = true;
                                                break;
                                            case TInterrupt.Action.STOP_LOOP:
                                                Logger.tracer($"Выход из цикла по условию: значение в ячейке x={cond.x} равно {cond.mustBe}");
                                                stopLoop = true;
                                                break;
                                        }
                                        continue;
                                    case TVariable var:
                                        SetVar(var, tmp[i, var.x]);
                                        break;
                                }
                            }
                        }
                    }

                    if (total > maxY - startY + 1)
                    {
                        Logger.warn("Попытка выйти за пределы документа, выход из цикла");
                        EOF = true;
                        break;
                    }

                    if (stopLoop)
                    {
                        Logger.debug("Выход из цикла по условию");
                        EOF = true;
                        break;
                    }
                    if (skipRecord) continue;
                    foreach (var var in dynamicVars.Values)
                        SetVar(var, tmp[i, var.x]);

                    scriptEveryRow?.Invoke();
                    if (skipRecord) continue;
                    if (stopLoop)
                    {
                        Logger.debug("Выход из цикла вызван JS хуком row");
                        EOF = true;
                        break;
                    }

                    callback(stepScope);
                }
                watch.Stop();
                Logger.debug($"Сегмент в {buffer} элементов (с {begin} по {end}) обработан за {watch.ElapsedMilliseconds} мс");
                guiCallback?.Invoke(total);

                begin += buffer;
                end += buffer;
            }
            watchTotal.Stop();
            Logger.debug("Времени всего: " + watchTotal.ElapsedMilliseconds);
            Logger.debug("Строк обработано: " + total);
            Logger.debug("Размер буффера:" + buffer);
            scriptAfter?.Invoke();
        }

        protected void SetVar(TVariable var, object value)
        {
            exception_var = var;
            var.Set(value);
            stepScope[var.name] = var;
        } 

        protected void FinalChecks()
        {
            int num = 1;

            if (validators == null) return;
            foreach (var validate in validators)
            {
                stepScope.TryGetValue(validate.var1, out TVariable var1);
                stepScope.TryGetValue(validate.var2, out TVariable var2);

                string value1 = var1?.value?.ToString() ?? "[неизвестно]";
                string value2 = var2?.value?.ToString() ?? "[неизвестно]";

                string elemMsg = validate.Message;
                string message;

                if (elemMsg == null)
                {
                    message = $"Финальная проверка №{num} провалена!";
                }
                else
                {
                    message = string.Format(elemMsg, value1, value2, num);
                    message = message.Replace("\\n", "\n");
                }

                if (var1 == null || var2 == null || var1.value == null || var2.value == null) throw new Exception(message);

                Logger.info($"Проверка номер {num} : {var1.name}({value1}) сравнивается с {var2.name}({value2})");

                bool isEqual;
                if (validate.Math != null)
                {
                    int count = validate.Math.count;
                    float prec = Single.Parse(validate.Math.precision);

                    float allowed_precision = prec / count * total;
                    float var1fl = Convert.ToSingle(var1.value);
                    float var2fl = Convert.ToSingle(var2.value);

                    Logger.info("var1 = " + var1fl.ToString("G9"));
                    Logger.info("var2 = " + var2fl.ToString("G9"));

                    if (Equals(var1fl, var2fl)) isEqual = true;
                    else
                    {
                        float diff = Math.Abs(Math.Abs(var1fl) - Math.Abs(var2fl));
                        isEqual = diff < allowed_precision;

                        string message_diff = string.Format(validate.Math.message, allowed_precision, diff).Replace("\\n", "\n");
                        message += "\n" + message_diff;
                        Logger.info(message_diff);
                    }
                }
                else isEqual = var1.value.Equals(var2.value);

                if (!isEqual) throw new Exception(message);

                num++;
            }
        }

        protected void InitVariables(Xml_Form lForm)
        {
            if (lForm.Fields.IF != null)
                foreach (var xmlelem in lForm.Fields.IF)
                {
                    XElement xelem = XElement.Parse(xmlelem.OuterXml);
                    conditions.Add(ScanCondition(xelem));
                }

            if (lForm.Fields.Static != null)
                foreach (var xmlelem in lForm.Fields.Static)
                {
                    XElement xelem = XElement.Parse(xmlelem.OuterXml);
                    AddVar(staticVars, getVar(xelem, false));
                }

            if (lForm.Fields.Dynamic != null)
                foreach (var xmlelem in lForm.Fields.Dynamic)
                {
                    XElement xelem = XElement.Parse(xmlelem.OuterXml);
                    AddVar(dynamicVars, getVar(xelem, true));
                }
        }

        protected TCondition ScanCondition(XElement xml)
        {
            string x = xml.Attribute("X")?.Value ??
                       throw new NullReferenceException("Attribute \"X\" can't be null!");
            string value = xml.Attribute("VALUE")?.Value ??
                           throw new NullReferenceException("Attribute \"VALUE\" can't be null!");
            bool isRegex = new[] { "1", "true", "yes"}.Contains(xml.Attribute("REGEX")?.Value);
            var xthen = xml.Element("THEN") ??
                        throw new NullReferenceException("Element <THEN> can't be null!");
            var xelse = xml.Element("ELSE");

            TCondition condition = new TCondition
            {
                x = int.Parse(x),
                mustBe = value,
                isRegex = isRegex
            };

            AddTActionsToList(condition.onTrue, xthen);
            AddTActionsToList(condition.onFalse, xelse);

            return condition;
        }

        protected void AddTActionsToList(IList<TAction> list, XElement target)
        {
            if (target == null) return;
            foreach (XElement elem in target.Elements())
            {
                TAction action = null;
                switch (elem.Name.ToString())
                {
                    case "SKIP_RECORD":
                        action = new TInterrupt(TInterrupt.Action.SKIP_RECORD);
                        break;
                    case "STOP_LOOP":
                        action = new TInterrupt(TInterrupt.Action.STOP_LOOP);
                        break;
                    case "Dynamic":
                        action = getVar(elem, true);
                        break;
                }
                if (action != null) list.Add(action);
            }
        }

        protected void AddVar(IDictionary<string, TVariable> dictionary, TVariable variable)
        {
            dictionary.Add(variable.name, variable);
        }

        protected TVariable getVar(XElement xml, bool dynamic)
        {
            var name = xml.Attribute("name")?.Value ?? throw new NullReferenceException("Variable attribute 'name' can't be null!");
            var ctype = xml.Attribute("type")?.Value ?? "string";

            TVariable variable;
            switch (ctype)
            {
                case "numeric":
                    variable = new TNumeric(name);
                    break;
                case "date":
                    variable = new TDate(name);
                    break;
                default:
                    variable = new TVariable(name);
                    break;
            }

            variable.x = Int32.Parse(xml.Attribute("X")?.Value ?? throw new NullReferenceException("Variable attribute 'X' can't be null!"));
            if (!dynamic) variable.y = Int32.Parse(xml.Attribute("Y")?.Value ?? throw new NullReferenceException("Variable attribute 'Y' can't be null!"));
            variable.dynamic = dynamic;

            if (variable is TNumeric tnumeric)
            {
                var function = xml.Attribute("function");

                if (function != null)
                    tnumeric.function = TNumeric.getFuncByString(function.Value);
            }

            if (variable is TDate tdate)
            {
                var lastday = xml.Attribute("lastday");
                var language = xml.Attribute("language");
                var format = xml.Attribute("format");

                if (lastday != null)
                    tdate.lastday = Boolean.Parse(lastday.Value);
                if (language != null)
                    tdate.language = language.Value;
                if (format != null)
                    tdate.format = format.Value;
            }

            variable.regex_replace = xml.Attribute("regex_replace")?.Value;

            var regex_pattern = xml.Attribute("regex_pattern");
            if (regex_pattern != null)
            {
                variable.regex_pattern = new Regex(regex_pattern.Value, RegexOptions.Compiled);

                var regex_group = xml.Attribute("regex_group");
                variable.regex_group = int.Parse(regex_group?.Value ?? "1");
            }
            return variable;
        }
    }

    public class MyException : Exception
    {
        public MyException(string message, Exception exp) : base(message, exp)
        {
            StackTrace = exp.StackTrace;
        }

        public override string StackTrace { get; }
    }

}
