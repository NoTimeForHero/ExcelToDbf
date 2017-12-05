using NickBuhro.Translit;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace DomofonExcelToDbf.Sources
{
    public class JS
    {
        public Jint.Engine engine;
        protected Regex regExS = new Regex(@"\s+", RegexOptions.Compiled);

        public delegate string DelegateReadExcel(int x, int y);
        public delegate void DelegateLog(object obj);

        /// <summary>
        /// Конструктор класса JS, реализующий все необходимые базовые функции
        /// 
        /// ---- Доступные функции: -----
        /// string translit(string input) - возвращает строку input в транслите
        /// string nospace(string input,string replaced) - заменяет в строке input все пробелы на replaced и возвращает строку
        /// string|null xls(int x, int y) - читает значение из ячейки Excel, возвращает null если произошла ошибка
        /// string|null afterRegEx(string input, Regex regex, int id=1) - разделяет строку input по регулярному выражению regex и возвращает id элемент полученного массива (1 если не указано) или null
        /// string|null dir(int id) - возвращает сегмент пути по заданному пути
        /// void log(string message) - вывести сообщение через Console.WriteLine (по умолчанию)
        /// void  string message) - кидает исключение класса Jint.Runtime.JavaScriptException с сообщением message
        /// 
        /// ---- Доступные переменные: ----
        /// string file - оригинальное имя Excel файла
        /// string dirCount - количество сегментов в пути
        /// 
        /// На выход должна подаваться единственная строка с новым именем файла
        /// </summary>
        public JS(DelegateReadExcel readExcel, DelegateLog log = null)
        {
            if (log == null) log = Console.WriteLine;

            engine = new Jint.Engine();
            engine.SetValue("translit", new Func<string, string>(FuncTranslit));
            engine.SetValue("nospace", new Func<string, string, string>(FuncReplaceSpace));
            engine.SetValue("afterRegEx", new Func<string, Regex, object, string>(FuncAfterRegEx));
            engine.SetValue("error", new Action<string>(FuncThrowException));
            engine.SetValue("log", log);
            engine.SetValue("xls", readExcel);
            engine.SetValue("dir", new Action(() => FuncThrowException("Ошибка 1754: Невозможно выполнить функцию dir(...), так как не установлена директория до конечного файла через JS->SetPath(...)!")));
        }

        public string Execute(string script)
        {
            return engine.Execute(script).GetCompletionValue().ToObject().ToString();
        }

        public void SetPath(string fullPath)
        {
            string dirPath = Path.GetDirectoryName(fullPath);
            if (dirPath == null) throw new InvalidOperationException($"Can't get directory name from: {fullPath}!");

            DirectoryInfo dir = new DirectoryInfo(dirPath);
            // В этом методе возможно утечка памяти, только непонятно как её устранить без разбиения на класс
            PathHelper helper = new PathHelper(dir);
            engine.SetValue("dir", new Func<int, string>(helper.GetLevel));
            engine.SetValue("file", Path.GetFileNameWithoutExtension(fullPath));
            engine.SetValue("dirCount", helper.Count);
            // Старые способы задания
            // Func<int, string> funcDir = (int level) => helper.GetLevel(level);
            // Func<int,string> funcDir = new Func<int,string>(helper.GetLevel);
        }

        protected string FuncTranslit(string input)
        {
            return SafeString(Transliteration.CyrillicToLatin(input, Language.Russian));
        }

        protected string SafeString(string result)
        {
            Array.ForEach(Path.GetInvalidFileNameChars(),
                  c => result = result.Replace(c.ToString(), String.Empty));
            return result;
        }

        protected string FuncReplaceSpace(string input, string replace)
        {
            return regExS.Replace(input, replace ?? "");
        }

        protected string FuncAfterRegEx(String input, Regex info, object nid)
        {
            int id = nid != null ? Convert.ToInt32(nid) : 1; // 1 == default
            string[] groups = info.Split(input);
            if (id > groups.Length - 1) return null;
            return groups[id];
        }

        protected void FuncThrowException(String text)
        {
            throw new Jint.Runtime.JavaScriptException("Исключение вызванное из JavaScript:\n" + text);
        }
    }
}
