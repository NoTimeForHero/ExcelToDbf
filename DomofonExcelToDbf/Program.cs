using Microsoft.Office.Interop.Excel;
using NickBuhro.Translit;
using SocialExplorer.IO.FastDBF;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DomofonExcelToDbf
{
    class DBF
    {
        public DbfFile odbf;
        public IEnumerable<XElement> dbfields;
        public int records = 0;
        public bool closed = false;
        protected string path;

        public DBF(String path, Encoding encoding = null)
        {
            this.path = path;
            // Если мы не передали кодировку, то используем DOS (=866)
            // Нельзя писать DBF(xxx, Encoding encoding = Encoding.GetEncoding(866)) так как аргументы метода должны вычисляться на этапе компиляции
            // А Encoding.GetEncoding(866) можно высчитать только при запуске приложения
            if (encoding == null) encoding = Encoding.GetEncoding(866);

            odbf = new DbfFile(encoding);
            odbf.Open(path, FileMode.Create); // FileMode.Create = файл будет перезаписан если уже существует
            Console.WriteLine("Создаём DBF с именем {0} и\nкодировкой: {1}", path, encoding);
        }

        // Эту функцию нельзя вызвать за пределами данного класса
        public void writeHeader(XElement form)
        {
            dbfields = form.Element("DBF").Elements("field");
            Console.WriteLine("Записываем в DBF {0} полей", dbfields.Count());
            foreach (XElement field in dbfields)
            {
                string input = field.Value;
                string name = field.Attribute("name").Value;
                string type = field.Attribute("type").Value;

                XAttribute attrlen = field.Attribute("length");

                DbfColumn.DbfColumnType column = DbfColumn.DbfColumnType.Character;
                if (type == "string") column = DbfColumn.DbfColumnType.Character;
                if (type == "date") column = DbfColumn.DbfColumnType.Date;
                if (type == "numeric") column = DbfColumn.DbfColumnType.Number;

                if (attrlen != null)
                {
                    var length = attrlen.Value.Split(',');
                    int nlen = Int32.Parse(length[0]);
                    int ndec = (length.Length > 1) ? Int32.Parse(length[1]) : 0;
                    odbf.Header.AddColumn(new DbfColumn(name, column, nlen, ndec));
                    Console.WriteLine("Записываем поле '{0}' типа '{1}' длиной {2},{3}", name, type, nlen, ndec);
                } else
                {
                    odbf.Header.AddColumn(new DbfColumn(name, column));
                    Console.WriteLine("Записываем поле '{0}' типа '{1}'", name, type);
                }
            }
            odbf.WriteHeader();    
        }

        public void appendRecord(Dictionary<string, object> variables)
        {
            var orec = new DbfRecord(odbf.Header);
            //orec.AllowIntegerTruncate = true;
            orec.AllowStringTurncate = true;

            int fid = 0;
            foreach (XElement field in dbfields)
            {

                string input = field.Value;
                string name = field.Attribute("name").Value;
                string type = XmlCondition.attrOrDefault(field, "type", "string");

                var matches = Regex.Matches(input, "\\$([0-9a-zA-Z]+)", RegexOptions.Compiled);
                foreach (Match m in matches)
                {
                    var repvar = m.Groups[1].Value;

                    if (!variables.ContainsKey(repvar)) // чтобы в финальном файле не оказалось строк вида $VARIABLE
                    {
                        input = input.Replace(m.Value, "");
                        continue;
                    }

                    object data = variables[repvar];
                    if (data == null) data = "";

                    if (type == "string" || type == "numeric")
                    {
                        input = input.Replace(m.Value, data.ToString());
                    }
                    else if (type == "date")
                    {
                        string format = XmlCondition.attrOrDefault(field, "format", "yyyy-MM-dd");
                        input = input.Replace(m.Value, ((DateTime)data).ToString(format));
                    }
                }

                orec[fid] = input;
                fid++;
            }

            odbf.Write(orec, true);
            //if (i < 20) foreach (var x in variables) Console.WriteLine(x.Key + "=" + x.Value);

            records++;
            if (records % 100 > 0) return;
            Console.WriteLine("Записей обработано: {0}", records);
        }

        public void close()
        {
            if (closed) return;
            closed = false;
            odbf.Close();
        }

        public void delete()
        {
            if (closed) return;
            close();

            File.Delete(this.path);
        }

    }

    class Excel
    {
        Microsoft.Office.Interop.Excel.Application app;
        Workbook wb;
        public Worksheet worksheet;
        protected bool saveMemory;

        public Excel(bool saveMemory)
        {
            if (saveMemory) app = new Microsoft.Office.Interop.Excel.Application();
            this.saveMemory = saveMemory;

        }

        public bool OpenWorksheet(String path)
        {
            // Если не экономим память, то создаём новый экземпляр COM OLE
            if (saveMemory)
            {
                if (wb != null) wb.Close(0);
            } else
            {
                if (app != null) app.Quit();
                app = new Microsoft.Office.Interop.Excel.Application();
            }

            wb = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            if (wb.Worksheets.Count < 1)
            {
                Console.WriteLine("Выбранный Excel не содержит ни одного листа!");
                return false;
            }

            worksheet = wb.Worksheets[1];
            return true;
        }

        public void close()
        {
            if (wb != null) wb.Close(0);
            if (app != null) app.Quit();
        }


    }

    class XmlCondition
    {
        public int x;
        public String value;

        public XElement then;
        public XElement or;

        public override string ToString()
        {
            String total = "";
            total += String.Format("X={0}",x) + "\n";
            total += String.Format("Value={0}",value) + "\n";
            total += String.Format("BEGIN:\n  {0}",then) + "\n";
            total += String.Format("ELSE:\n  {0}", or) + "\n\n";
            return total;
        }

        public static string attrOrDefault(XElement element, String attr, String def)
        {
            XAttribute xattr = element.Attribute(attr);
            if (xattr == null) return def;
            return xattr.Value;
        }

        public static List<XmlCondition> makeList(XElement form)
        {
            var conditions = new List<XmlCondition>();

            var local = form.Element("Fields").Elements("IF");
            foreach (XElement elem in local)
            {
                var cond = new XmlCondition();

                cond.x = Int32.Parse(elem.Attribute("X").Value);
                cond.value = elem.Value;

                // Получаем секцию THEN, так как она обязана быть следующей после IF
                cond.then = (XElement)elem.NextNode;

                // А вот секции ELSE может и не быть
                var next = elem.NextNode.NextNode;

                var nextName = ((XElement)next).Name.ToString();
                if (nextName == "ELSE") cond.or = (XElement)next;

                conditions.Add(cond);
            }
            return conditions;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            new Program();
        }

        public Program()
        {
            String confName = Path.ChangeExtension(System.AppDomain.CurrentDomain.FriendlyName, ".xml");

            if (!File.Exists(confName) || true)
            {
                Console.WriteLine("Не найден конфигурационный файл!");
                Console.WriteLine("Распаковываем его из внутренних ресурсов...");
                Tools.WriteResourceToFile("xConfig", confName);
            }

            XDocument xdoc = XDocument.Load(confName);

            String dirInput = Tools.getDirectory(xdoc, "inputDirectory"); 
            String dirOutput = Tools.getDirectory(xdoc, "outputDirectory");

            Console.WriteLine("Директория чтения: {0}", dirInput);
            Console.WriteLine("Директория записи: {0}", dirOutput);

            bool onlyRules = xdoc.Root.Element("only_rules").Value == "true";
            bool saveMemory = xdoc.Root.Element("save_memory").Value == "true"; ; // экономить память, если включено то будет использоваться один инстанс COM Excel с переключением Worksheet

            var formToFile = new Dictionary<string, string>();
            var outlog = new List<string>();
            var files = new HashSet<string>();
            foreach (var extension in xdoc.Root.Element("extensions").Elements("ext"))
            {
                string []fbyext = Directory.GetFiles(dirInput, extension.Value, SearchOption.TopDirectoryOnly);
                fbyext = fbyext.Where(path => !Path.GetFileName(path).StartsWith("~$")).ToArray(); // Игнорируем временные файлы Excel вида ~$Document.xls[x]
                files.UnionWith(fbyext);
                Console.WriteLine("Файлов найдено {1} по маске {0}", extension.Value, fbyext.Length);
            }

            Excel excel = new Excel(saveMemory);
            DBF dbf = null;

            var totalwatch = new System.Diagnostics.Stopwatch();
            totalwatch.Start();
            foreach (string fname in files)
            {
                for (int i = 0; i < 2; i++) Console.WriteLine();

                // COM Excel требуется полный путь до файла
                string finput = Path.GetFullPath(fname);

                bool deleteDbf = false;

                try
                {
                    Console.WriteLine("Загружаем Excel документ: {0}", Path.GetFileName(finput));
                    excel.OpenWorksheet(finput);

                    var form = Tools.findCorrectForm(excel.worksheet, xdoc);
                    string foutput = Path.Combine(dirOutput, Tools.getOutputFilename(excel.worksheet, xdoc, dirInput, finput));

                    if (onlyRules)
                    {
                        var formname = (form == null) ? "null" : form.Element("Name").Value;
                        formToFile.Add(Path.GetFileName(finput), formname);
                        continue;
                    }

                    if (form == null)
                    {
                        Console.WriteLine("Не найдено подходящих форм для обработки документа work.xml!");
                        continue;
                    }

                    dbf = new DBF(foutput);
                    dbf.writeHeader(form);

                    var stopwatch = new System.Diagnostics.Stopwatch();

                    stopwatch.Start();
                    Tools.eachRecord(excel.worksheet, form, dbf.appendRecord);
                    stopwatch.Stop();

                    Console.WriteLine("Времени потрачено на обработку данных: {0}", stopwatch.Elapsed);
                    Console.WriteLine("Обработано записей: {0} ", dbf.records);
                    outlog.Add(String.Format("Файл {0} в {1} записей обработан за {2}",Path.GetFileName(finput),dbf.records,stopwatch.Elapsed));

                    int startY = Tools.startY(form);
                    Console.WriteLine("Начиная с {0} по {1}", startY, startY + dbf.records);
                }
                catch (Exception ex)
                {
                    deleteDbf = true;
                    Console.Error.WriteLine(ex);
                }
                finally
                {
                    Console.WriteLine("Закрытие COM Excel и DBF");
                    if (dbf != null) dbf.close();
                    if (dbf != null && deleteDbf) dbf.delete();
                }

            }
            totalwatch.Stop();

            // Не забываем завершить Excel
            excel.close();

            if (onlyRules)
            {
                for (int i = 0; i < 3; i++) Console.WriteLine();
                foreach (var tup in formToFile)
                {
                    Console.WriteLine("Для файла {0} выбрана форма {1}", tup.Key, tup.Value);
                }
            }

            foreach (string line in outlog) Console.WriteLine(line);

            Console.WriteLine("Времени затрачено суммарно: {0}", totalwatch.Elapsed);

            for (int i = 0; i < 2; i++) Console.WriteLine();

            Console.WriteLine("Нажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
    
    class Tools {         

        public static string getDirectory(XDocument xdoc, String type)
        {
            var elem = xdoc.Root.Element(type);
            var regex = elem.Attribute("regex");

            if (regex != null)
            {
                var lastDir = getDirectoryName(elem.Value); // Проверяем только последнюю директорию во всём пути
                Match match = (new Regex(regex.Value)).Match(lastDir);
                if (!match.Success) throw new ArgumentException(String.Format("Директория '{0}' не попадает под регулярное выражение '{1}'!",lastDir, regex.Value));
            }

            return elem.Value;
        }

        public static string getOutputFilename(Worksheet worksheet, XDocument xdoc, String inputDirectory, String inputFile)
        {
            XElement outfile = xdoc.Root.Element("outfile");

            bool simple = outfile.Element("simple").Value == "true";
            if (simple) return Path.GetFileName(Path.ChangeExtension(inputFile, ".dbf"));

            var x = Int32.Parse(outfile.Element("X").Value);
            var y = Int32.Parse(outfile.Element("Y").Value);

            string cAfter = outfile.Element("after").Value;
            string fullName = worksheet.Cells[y, x].Value;

            int nAfter = fullName.IndexOf(cAfter);
            if (nAfter < 0) throw new ArgumentNullException(String.Format("Подстрока '{0}' не найдена в строке '{1}'!",cAfter,fullName));

            string regionName = fullName.Substring(nAfter + cAfter.Length);

            // Транслит если нужно
            bool translit = outfile.Element("translit").Value == "true";
            if (translit) regionName = Transliteration.CyrillicToLatin(regionName, Language.Russian);

            // Заменяем пробелы в имени файла на заданный в конфиге символ/подстроку
            string replaceSpaceWith = outfile.Element("spaces").Value;
            regionName = regionName.Replace(" ", replaceSpaceWith);

            // Нужно ли добавлять имя директории перед файлом
            bool dirname = outfile.Element("include_dir_name").Value == "true";
            if (dirname)
            {
                string delim = outfile.Element("dir_delimiter").Value;
                regionName = getDirectoryName(inputDirectory) + delim + regionName;
            }

            // Не забываем добавить расширение на конец
            regionName = regionName + ".dbf";

            return regionName;
        }

        public static string getDirectoryName(String path)
        {
            if (Path.GetExtension(path) == "") return Path.GetFileName(path);
            return new FileInfo(path).Directory.Name;
        }

        public static Int32 startY(XElement form)
        {
            return Int32.Parse(form.Element("Fields").Element("StartY").Value);
        }

        public static void eachRecord(Worksheet worksheet, XElement form, Action<Dictionary<string,object>> callback)
        {
            Dictionary<string, object> variables = new Dictionary<string, object>();

            // Позиция с которой начинаются данные
            var minY = startY(form);
            var maxY = worksheet.UsedRange.Rows.Count;

            // Получаем список статических переменных, которые не меняются для всех записей в данном листе
            var staticvars = form.Element("Fields").Elements("Static");
            foreach (XElement staticvar in staticvars)
            {
                var x = Int32.Parse(staticvar.Attribute("X").Value);
                var y = Int32.Parse(staticvar.Attribute("Y").Value);

                var name = staticvar.Attribute("name").Value;

                var cell = worksheet.Cells[y, x].Value;
                variables.Add(name, getVar(staticvar, cell));
            }

            var dynamicvars = form.Element("Fields").Elements("Dynamic");
            var conditions = XmlCondition.makeList(form);
            // Начинаем обходить каждый лист
            for (int y = minY; y < maxY; y++)
            {
                // Получаем значения динамических переменных без условий
                foreach (XElement dyvar in dynamicvars)
                {
                    var x = Int32.Parse(dyvar.Attribute("X").Value);
                    var name = dyvar.Attribute("name").Value;

                    var cell = worksheet.Cells[y, x].Value;
                    variables[name] = getVar(dyvar, cell);
                }

                // Проверяем каждое условие
                foreach (XmlCondition cond in conditions)
                {
                    var cell = worksheet.Cells[y, cond.x].Text;

                    XElement section = (cell == cond.value) ? cond.then : cond.or;
                    if (section == null) continue;

                    var condvars = section.Elements("Dynamic");
                    foreach (XElement dyvar in condvars)
                    {
                        var x = Int32.Parse(dyvar.Attribute("X").Value);
                        var name = dyvar.Attribute("name").Value;

                        try
                        {
                            cell = worksheet.Cells[y, x].Value;
                            variables[name] = getVar(dyvar, cell);
                        } catch (Exception)
                        {
                            Console.WriteLine("Ошибка в переменной {0} на Y={1},X={2}", name, y, x);
                            throw;
                        }
                    }

                    if (section.Element("SKIP_RECORD") != null)
                    {
                        Console.WriteLine("Пропускаем строку Y={0}", y);
                        goto skip_record;
                    }
                    if (section.Element("STOP_LOOP") != null)
                    {
                        Console.WriteLine("Выходим из цикла на Y={0} по условию X[{1}]={2}", y, cond.x, cond.value);
                        goto skip_loop;
                    }
                }

                callback(variables);
                skip_record:;
            }
            skip_loop:;

            Console.WriteLine("Составление записей завершено?");

        }

        // <summary>
        // Метод считывает внутренний ресурс и записывает его в файл, возвращая статус существования ресурса
        // </summary>
        // <param name="var">Имя внутренного ресурса</param>
        // <param name="cell">Имя внутренного ресурса</param>
        // <returns>false если внутренний ресурс не был найден</returns>
        public static object getVar(XElement var, object obj)
        {
            if (obj == null)
            {
                return null;
            }

            String type = XmlCondition.attrOrDefault(var, "type", "string");
            String cell = obj.ToString();

            if (type == "string")
            {
                return cell;
            }
            if (type == "numeric")
            {
                return Double.Parse(cell);
            }

            if (type == "date") {
                var format = var.Attribute("format").Value;
                var language = XmlCondition.attrOrDefault(var, "language", "ru-ru");
                DateTime date = DateTime.ParseExact(cell, format, CultureInfo.GetCultureInfo(language));

                // Если нам нужен последний день в месяце
                string lastday = XmlCondition.attrOrDefault(var, "lastday", "");
                if (lastday == "true") date = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
                return date;
            }

            return null;
        }


        // <summary>
        // Ищет подходящую XML форму для документа или null если ни одна не подходит
        // </summary>
        public static XElement findCorrectForm(Worksheet worksheet, XDocument xdoc)
        {
            var forms = xdoc.Root.Element("Forms").Elements("Form").ToList();

            foreach (XElement form in forms)
            {
                bool correct = true;
                String name = form.Element("Name").Value;
                Console.WriteLine(String.Format("Проверяем форму \"{0}\"",name));

                var equals = form.Element("Rules").Elements("Equal");
                foreach (XElement equal in equals)
                {
                    var x = Int32.Parse(equal.Attribute("X").Value);
                    var y = Int32.Parse(equal.Attribute("Y").Value);
                    var mustbe = equal.Value;

                    string cell = null;

                    try
                    {
                        cell = worksheet.Cells[y, x].Value.ToString();
                    } catch (Exception ex)
                    {
                        Console.WriteLine(String.Format("Произошла ошибка при чтении ячейки Y={0},X={1}!", y, x));
                        Console.WriteLine(String.Format("Ожидалось: {0}", mustbe));
                        Console.WriteLine("Ошибка: {0}",ex.Message);
                        correct = false;
                        break;
                    }

                    if (mustbe != cell)
                        {
                            Console.WriteLine(String.Format("Проверка провалена (Y={0},X={1})",y,x));
                            Console.WriteLine(String.Format("Ожидалось: {0}", mustbe));
                            Console.WriteLine(String.Format("Найдено: {0}", cell));
                            correct = false;
                            break;
                        }
                        Console.WriteLine(String.Format("Y={0},X={1}:  {2}=={3}",y,x,mustbe,cell));
                    }
                    if (correct) return form;
            }
            return null;
        }

        // <summary>
        // Метод считывает внутренний ресурс и записывает его в файл, возвращая статус существования ресурса
        // </summary>
        // <param name="resourceName">Имя внутренного ресурса</param>
        // <param name="resourceName">Имя внутренного ресурса</param>
        // <returns>false если внутренний ресурс не был найден</returns>
        public static bool WriteResourceToFile(string resourceName, string fileName)
        {
            using (var resource = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (resource == null) return false;
                using (var file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    resource.CopyTo(file);
                }
            }
            return true;
        }


    }
}
