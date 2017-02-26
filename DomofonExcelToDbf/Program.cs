using Microsoft.Office.Interop.Excel;
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

        public DBF(String path, Encoding encoding = null)
        {
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

        public void append(String kp, String fio, String adres, String summa, String dateopl)
        {
            var orec = new DbfRecord(odbf.Header);
            orec.AllowDecimalTruncate = true; // Отсекает дробную часть числа?
            orec[0] = kp;
            orec[1] = fio;
            orec[2] = adres;
            orec[3] = summa;
            orec[4] = dateopl;
            odbf.Write(orec, true); 
        }

        public void close()
        {
            odbf.Close();
        }

    }

    class Excel
    {
        Microsoft.Office.Interop.Excel.Application app;
        Workbook wb;
        public Worksheet worksheet;

        public Excel(String path)
        {
            app = new Microsoft.Office.Interop.Excel.Application();

            wb = app.Workbooks.Open(@"C:\Test\work.xml", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

            if (wb.Worksheets.Count < 1)
            {
                Console.WriteLine("Выбранный Excel не содержит ни одного листа!");
            }
        
            worksheet = wb.Worksheets[1];
        }

        public void close()
        {
            wb.Close(0);
            app.Quit();
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
            Tools.WriteResourceToFile("xConfig", "config.xml");
            XDocument xdoc = XDocument.Load("config.xml");

            String dirInput = xdoc.Root.Element("inputDirectory").Value;
            String dirOutput = xdoc.Root.Element("outputDirectory").Value;

            Console.WriteLine("Директория чтения: {0}", dirInput);
            Console.WriteLine("Директория записи: {0}", dirOutput);

            var files = new HashSet<string>();
            foreach (var extension in xdoc.Root.Element("extensions").Elements("ext"))
            {
                string []fbyext = Directory.GetFiles(dirInput, extension.Value, SearchOption.TopDirectoryOnly);
                files.UnionWith(fbyext);
                Console.WriteLine("Файлов найдено {1} по маске {0}", extension.Value, fbyext.Length);
            }

            Excel excel = null;
            DBF dbf = null;

            foreach (string fname in files)
            {
                for (int i = 0; i < 2; i++) Console.WriteLine();

                // COM Excel требуется полный путь до файла
                string finput = Path.GetFullPath(fname);
                var foutput = Path.Combine(dirOutput, Path.GetFileName(Path.ChangeExtension(finput, ".dbf")));

                try
                {
                    Console.WriteLine("Загружаем Excel документ: {0}", Path.GetFileName(finput));
                    excel = new Excel(finput);

                    var form = Tools.findCorrectForm(excel.worksheet, xdoc);
                    if (form == null)
                    {
                        Console.WriteLine("Не найдено подходящих форм для обработки документа work.xml!");
                        break;
                    }

                    dbf = new DBF(foutput);
                    dbf.writeHeader(form);

                    var stopwatch = new System.Diagnostics.Stopwatch();

                    stopwatch.Start();
                    Tools.eachRecord(excel.worksheet, form, dbf.appendRecord);
                    stopwatch.Stop();

                    Console.WriteLine("Времени потрачено на обработку данных: {0}", stopwatch.Elapsed);
                    Console.WriteLine("Обработано записей: {0} ", dbf.records);

                    int startY = Tools.startY(form);
                    Console.WriteLine("Начиная с {0} по {1}", startY, startY + dbf.records);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex);
                }
                finally
                {
                    Console.WriteLine("Закрытие COM Excel и DBF");
                    if (dbf != null) dbf.close();
                    if (excel != null) excel.close();
                }
            }

            for (int i = 0; i < 2; i++) Console.WriteLine();
            Console.WriteLine("End?");

        }
    }
    
    class Tools {         

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

                        cell = worksheet.Cells[y, x].Value;
                        variables[name] = getVar(dyvar, cell);
                    }

                    if (section.Element("SKIP_RECORD") != null)
                        goto skip_record;
                    if (section.Element("STOP_LOOP") != null)
                        goto skip_loop;
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

            if (type == "string" || type == "numeric")
            {
                return cell;
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
                        Console.WriteLine(String.Format("X={0},Y={1}:  {2}=={3}",x,y,mustbe,cell));
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
