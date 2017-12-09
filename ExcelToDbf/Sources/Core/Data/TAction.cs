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
        public Type type;

        public bool dynamic;
        public int x;
        public int y;

        public object value;

        public bool use_regex = false;
        public Regex regex_pattern;
        public int regex_group;

        public TVariable(string name)
        {
            this.name = name ?? throw new ArgumentNullException(nameof(name), @"Name can't be null!");
        }

        public void Set(object val)
        {
            string str = val?.ToString() ?? "";
            if (use_regex)
                str = MatchGroup(str, regex_pattern, regex_group);

            switch (this)
            {
                case TDate tdate:
                    tdate.Set(str);
                    break;
                case TNumeric tnum:
                    tnum.Set(str);
                    break;
                default:
                    value = str;
                    break;
            }
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

        public static Type getByString(string str)
        {
            switch (str)
            {
                case "string":
                    return Type.EString;
                case "date":
                    return Type.EDate;
                case "numeric":
                    return Type.ENumeric;
                default:
                    return Type.EUnknown;
            }
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

        public new void Set(object obj)
        {
            if ("".Equals(obj)) obj = "0"; // Иначе Convert.ToSingle упадёт с ошибкой
            float flValue = Convert.ToSingle(obj);
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

        public new void Set(object val)
        {
            DateTime date = DateTime.ParseExact(val.ToString(), format, CultureInfo.GetCultureInfo(language));
            if (lastday) date = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            value = date;
        }
    }
}
