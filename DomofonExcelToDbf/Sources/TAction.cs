using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DomofonExcelToDbf.Sources
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
        public enum Type : byte
        {
            EUnknown,
            EString,
            ENumeric,
            EDate
        }

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
            this.name = name;
        }

        public void Set(object val)
        {
            string str = (val == null) ? "" : val.ToString();
            if (use_regex)
                str = RegExCache.MatchGroup(str, regex_pattern, regex_group);

            if (false) ;
            else if (this is TDate tdate) tdate.Set(str);
            else if (this is TNumeric tnum) tnum.Set(str);
            else this.value = str;
        }

        public static Type getByString(string str)
        {
            if (str == "string") return Type.EString;
            if (str == "date") return Type.EDate;
            if (str == "numeric") return Type.ENumeric;
            return Type.EUnknown;
        }

        public override bool Equals(object obj)
        {
            var item = obj as TVariable;
            if (item == null) return false;
            return this.name == item.name;
        }

        public override int GetHashCode()
        {
            return name.GetHashCode();
        }
    }

    public class TNumeric : TVariable
    {
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

        public TNumeric(string name) : base(name) { }

        public new void Set(object obj)
        {
            if ("".Equals(obj)) obj = "0"; // Иначе Convert.ToSingle упадёт с ошибкой
            float value = Convert.ToSingle(obj);
            switch (function)
            {
                case Func.SUM:
                    this.value = Convert.ToSingle(this.value) + value;
                    break;
                default:
                    this.value = value;
                    break;
            }
        }
    }

    public class TDate : TVariable
    {
        public bool lastday = false;
        public string format = "dd.MM.yyy";
        public string language = "ru-ru";

        public TDate(string name) : base(name) { }

        public new void Set(object val)
        {
            DateTime date = DateTime.ParseExact(val as string, format, CultureInfo.GetCultureInfo(language));
            if (lastday) date = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            this.value = date;
        }
    }
}
