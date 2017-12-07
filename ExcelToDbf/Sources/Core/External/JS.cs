using System;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.CSharp.RuntimeBinder;
using NickBuhro.Translit;

namespace ExcelToDbf.Sources.Core.External
{
    public class JS
    {
        protected Jint.Engine engine;
        protected Regex regExS = new Regex(@"\s+", RegexOptions.Compiled);

        /// <summary>
        /// Функция чтения указанной ячейки Excel
        /// </summary>
        public delegate string DelegateReadExcel(int x, int y);

        /// <summary>
        /// Функция, логирующие данные из JS скрипта в основную программу
        /// </summary>
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

        /// <summary>
        /// Выполняет указанный скрипт и возвращает конечное имя файла
        /// </summary>
        public string Execute(string script)
        {
            return engine.Execute(script).GetCompletionValue().ToObject().ToString();
        }

        /// <summary>
        /// Задаёт JS движку функции и переменные, необходимые для работы с путём к файлу
        /// </summary>
        /// <param name="fullPath">Полный путь до файла, включая его имя</param>
        public JS SetPath(string fullPath)
        {
            string dirPath = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrEmpty(dirPath)) throw new ArgumentException($"Can't get directory name from: {fullPath}!");

            DirectoryInfo dir = new DirectoryInfo(dirPath);
            // В этом методе возможно утечка памяти, только непонятно как её устранить без разбиения на класс
            PathHelper helper = new PathHelper(dir);
            engine.SetValue("dir", new Func<int, string>(helper.GetLevel));
            engine.SetValue("file", Path.GetFileNameWithoutExtension(fullPath));
            engine.SetValue("dirCount", helper.Count);
            // Старые способы задания
            // Func<int, string> funcDir = (int level) => helper.GetLevel(level);
            // Func<int,string> funcDir = new Func<int,string>(helper.GetLevel);
            return this;
        }

        /// <summary>
        /// Переводит строку в транслит
        /// </summary>
        protected string FuncTranslit(string input)
        {
            return SafeString(Transliteration.CyrillicToLatin(input, Language.Russian));
        }

        /// <summary>
        /// Удаляет из строки все недопустимые для файловой системы символы
        /// </summary>
        protected string SafeString(string result)
        {
            Array.ForEach(Path.GetInvalidFileNameChars(),
                  c => result = result.Replace(c.ToString(), String.Empty));
            return result;
        }

        /// <summary>
        /// Заменяет все пробельные символы в строке на указанную строку
        /// </summary>
        protected string FuncReplaceSpace(string input, string replace)
        {
            return regExS.Replace(input, replace ?? "");
        }

        /// <summary>
        /// Разбивает подстроку input по регулярному выражению  info и возвращает nid группу
        /// Например: для построки abc с регуляркой one(two)(three)(four) и nid=2 вернёт "three"
        /// </summary>
        protected string FuncAfterRegEx(String input, Regex info, object nid)
        {
            int id = nid != null ? Convert.ToInt32(nid) : 1; // 1 == default
            string[] groups = info.Split(input);
            if (id > groups.Length - 1) return null;
            return groups[id];
        }

        /// <summary>
        /// Бросает исключение с заданным сообщением
        /// </summary>
        protected void FuncThrowException(String text)
        {
            throw new JSException("Исключение вызванное из JavaScript:\n" + text);
        }

        public class JSException : Exception
        {
            public JSException(string message) : base(message) { }
        }
    }
}
