using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace DomofonExcelToDbf.Sources.Core
{
    /// <summary>
    /// Класс для удобного разбития пути на сегменты
    /// Например C:\One\Two\Three превратятся в массив: "C:\", "One", "Two", "Three"
    /// Любой элемент из которого можно получить через метод GetLevel(index)
    /// </summary>
    public class PathHelper
    {

        public List<string> components;

        public int Count => components.Count;

        public PathHelper(DirectoryInfo input)
        {
            components = Split(input);
        }

        public string GetLevel(int index)
        {
            int last = components.Count - 1;
            if (last < index)
                return null;
            return components[last - index];
        }

        protected List<string> Split(DirectoryInfo path)
        {
            if (path == null) throw new ArgumentNullException(nameof(path));
            var ret = new List<string>();
            if (path.Parent != null) ret.AddRange(Split(path.Parent));
            ret.Add(path.Name);
            return ret;
        }
    }

    class RegExCache
    {
        protected Dictionary<String, Regex> regexes = new Dictionary<String, Regex>();

        protected Regex Prepare(String strregex)
        {
            if (!regexes.ContainsKey(strregex)) regexes.Add(strregex, new Regex(strregex, RegexOptions.IgnoreCase | RegexOptions.Compiled));
            return regexes[strregex];
        }

        public String Replace(String input, String strregex, String replacement = "$1")
        {
            Regex regex = Prepare(strregex);
            return regex.Replace(input, replacement);
        }

        public bool IsMatch(String input, String strregex)
        {
            Regex regex = Prepare(strregex);
            return regex.Match(input).Success;
        }

        public String MatchGroup(String input, String strregex, int group = 1)
        {
            Regex regex = Prepare(strregex);
            Match match = regex.Match(input);
            if (!match.Success) return "";
            if (match.Groups.Count - 1 < group) return "";
            return match.Groups[group].Value;
        }

    }

}
