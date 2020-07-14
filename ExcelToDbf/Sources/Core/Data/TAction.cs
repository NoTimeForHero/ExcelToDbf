using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelToDbf.Sources.Core.Data.TData
{
    /// <summary>
    /// Универсальный класс, от которого наследуются всё возможные операции
    /// Сюда входят: условия, переменные, прерывания цикла обработки
    /// </summary>
    public abstract class TAction { }

    public class TInterrupt : TAction
    {
        public Action action;

        public TInterrupt(Action action)
        {
            this.action = action;
        }

        public enum Action
        {
            SKIP_RECORD,
            STOP_LOOP
        }
    }

    public class TCondition : TAction
    {
        public int x;
        public string mustBe;

        public List<TAction> onTrue = new List<TAction>();
        public List<TAction> onFalse = new List<TAction>();
    }

    public class TVariable : TAction
    {
        public readonly string name;

        public bool dynamic;
        public int x;
        public int y;

        public object value;

        public Regex regex_pattern;
        public String regex_replace;
        public int regex_group = 1;

        public TVariable(string name)
        {
            this.name = name ?? throw new ArgumentNullException(nameof(name), @"Name can't be null!");
        }

        public virtual void Set(object val)
        {
            string str = ToStr(val);
            str = RegExProcess(str);
            value = str;
        }

        protected string ToStr(object val)
        {
            return val?.ToString() ?? "";
        }

        protected string RegExProcess(String str)
        {
            if (regex_replace != null) return regex_pattern.Replace(str, regex_replace);
            if (regex_pattern != null) return MatchGroup(str, regex_pattern, regex_group);
            return str;
        }

        protected static String MatchGroup(String input, Regex regex, int group = 1)
        {
            Match match = regex.Match(input);
            if (!match.Success) return "";
            if (match.Groups.Count - 1 < group) return "";
            return match.Groups[group].Value;
        }

        #region Class Enum

        public enum Type : byte
        {
            EUnknown,
            EString,
            ENumeric,
            EDate
        }

        #endregion

        #region Override Object Method's

        public override bool Equals(object obj)
        {
            if (!(obj is TVariable item)) return false;
            return name == item.name;
        }

        public override int GetHashCode()
        {
            return name.GetHashCode();
        }

        #endregion
    }

    public class TNumeric : TVariable
    {
        public TNumeric(string name) : base(name) { }

        public override void Set(object obj)
        {
            string str = ToStr(obj);

            str = RegExProcess(str);
            if (str == "") str = "0"; // Иначе Convert.ToSingle упадёт с ошибкой

            float flValue = Convert.ToSingle(str);

            switch (function)
            {
                case Func.SUM:
                    value = Convert.ToSingle(value) + flValue;
                    break;
                default:
                    value = flValue;
                    break;
            }
        }

        #region Class Enum

        public Func function = Func.NONE;

        public enum Func : byte
        {
            NONE,
            SUM
        }

        public static Func getFuncByString(string str)
        {
            if (str == "SUM") return Func.SUM;
            return Func.NONE;
        }

        #endregion

    }

    public class TDate : TVariable
    {
        public bool lastday = false;
        public string format = "dd.MM.yyyy";
        public string language = "ru-ru";

        public TDate(string name) : base(name) { }

        public override void Set(object val)
        {
            // Регулярные выражения игнорируются, если в ячейке уже дата
            if (val is DateTime xdate)
            {
                if (lastday) xdate = new DateTime(xdate.Year, xdate.Month, DateTime.DaysInMonth(xdate.Year, xdate.Month));
                value = xdate;
                return;
            }

            string str = ToStr(val);

            str = RegExProcess(str);

            DateTime date;
            try
            {
                date = DateTime.ParseExact(str, format, CultureInfo.GetCultureInfo(language));
            }
            catch (FormatException ex)
            {
                throw new FormatException($"Не удалось распознать строку \"{str}\" как валидную дату формата \"{format}\"!", ex);
            }

            if (lastday) date = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            value = date;
        }
    }
}
