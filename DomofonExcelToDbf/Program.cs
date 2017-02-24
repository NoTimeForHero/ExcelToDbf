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
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DomofonExcelToDbf
{
    class DBF
    {
        public DbfFile odbf;

        public DBF(String path, Encoding encoding = null)
        {
            // Если мы не передали кодировку, то используем DOS (=866)
            // Нельзя писать DBF(xxx, Encoding encoding = Encoding.GetEncoding(866)) так как аргументы метода должны вычисляться на этапе компиляции
            // А Encoding.GetEncoding(866) можно высчитать только при запуске приложения
            if (encoding == null) encoding = Encoding.GetEncoding(866);

            odbf = new DbfFile(encoding);
            odbf.Open(path, FileMode.Create); // FileMode.Create = файл будет перезаписан если уже существует
            writeHeader();
        }

        // Эту функцию нельзя вызвать за пределами данного класса
        protected void writeHeader()
        {
            odbf.Header.AddColumn(new DbfColumn("KP", DbfColumn.DbfColumnType.Character, 14, 0));
            odbf.Header.AddColumn(new DbfColumn("FIO", DbfColumn.DbfColumnType.Character, 40, 0));
            odbf.Header.AddColumn(new DbfColumn("ADRES", DbfColumn.DbfColumnType.Character, 40, 0));
            odbf.Header.AddColumn(new DbfColumn("SUMMA", DbfColumn.DbfColumnType.Number, 10, 2));
            odbf.Header.AddColumn(new DbfColumn("DATEOPL", DbfColumn.DbfColumnType.Date));
            odbf.WriteHeader();
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
            WriteResourceToFile("xConfig", "config.xml");
            XDocument xdoc = XDocument.Load("config.xml");

            var sheet = getWorksheet(@"C:\Test\work.xml");

            var forms = xdoc.Root.Element("Forms").Elements("Form").ToList();

            var form = findCorrectForm(sheet, forms);
            if (form == null)
            {
                Console.WriteLine("Не найдено подходящих форм для обработки документа work.xml!");
                return;
            }

            eachRecord(sheet, form, debugRecords);


            var a = DateTime.ParseExact("Декабря 2017", "MMMM yyyy", CultureInfo.GetCultureInfo("ru-ru"));
            //var b = DateTime.ParseExact("Декабрь 2017", "MMMM yyyy", CultureInfo.GetCultureInfo("ru-ru"));
            //Console.WriteLine(a);
            //Console.WriteLine(b);
            //createDBF();
            //readExcel();
            //readCOMExcel();
            //WriteResourceToFile("xConfig", "config.xml");
        }

        static int i = 0;
        public void debugRecords(Dictionary<string, object> variables)
        {
            foreach (var x in variables) Console.WriteLine(x.Key + "=" + x.Value);

            i++;
            if (i % 10 > 0) return;
            Console.WriteLine("Записей обработано: {0}",i);
        }

        private static void Benchmark(System.Action act, int iterations)
        {
            GC.Collect();
            act.Invoke(); // run once outside of loop to avoid initialization costs
            var sw = System.Diagnostics.Stopwatch.StartNew();
            for (int i = 0; i < iterations; i++)
            {
                act.Invoke();
            }
            sw.Stop();
            Console.WriteLine((sw.ElapsedMilliseconds / iterations).ToString());
        }

        public void eachRecord(Worksheet worksheet, XElement form, Action<Dictionary<string,object>> callback)
        {
            Dictionary<string, object> variables = new Dictionary<string, object>();

            // Позиция с которой начинаются данные
            var minY = Int32.Parse(form.Element("Fields").Element("StartY").Value);
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

                    if (section.Element("SKIP_RECORD") != null) goto skip_record;
                    if (section.Element("STOP_LOOP") != null) goto skip_loop;

                    var condvars = section.Elements("Dynamic");
                    foreach (XElement dyvar in condvars)
                    {
                        var x = Int32.Parse(dyvar.Attribute("X").Value);
                        var name = dyvar.Attribute("name").Value;

                        cell = worksheet.Cells[y, x].Value;
                        variables[name] = getVar(dyvar, cell);
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
        public object getVar(XElement var, object obj)
        {
            if (obj == null)
            {
                return null;
            }

            String type = attrOrDefault(var, "type", "string");
            String cell = obj.ToString();

            if (type == "string")
            {
                return cell;
            }

            if (type == "date") {
                var format = var.Attribute("format").Value;
                var language = attrOrDefault(var, "language", "ru-ru");
                DateTime date = DateTime.ParseExact(cell, format, CultureInfo.GetCultureInfo(language));

                // Если нам нужен последний день в месяце
                string lastday = attrOrDefault(var, "lastday", "");
                if (lastday == "true") date = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
                return date;
            }

            return null;
        }

        public string attrOrDefault(XElement element, String attr, String def) {
            XAttribute xattr = element.Attribute(attr);
            if (xattr == null) return def;
            return xattr.Value;
        }

        // <summary>
        // Ищет подходящую XML форму для документа или null если ни одна не подходит
        // </summary>
        public XElement findCorrectForm(Worksheet worksheet, List<XElement> forms)
        {
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
        public bool WriteResourceToFile(string resourceName, string fileName)
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

        private void createDBF()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), "TestNew2.dbf");
            DBF dbf = new DBF(path);
            dbf.append("22", "Иванов Иван Иванович", "Москва, Красная Площадь", "5555", "2017-01-01");
            dbf.close();
        }

        private Worksheet getWorksheet(string filepath)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();

            Workbook wb = app.Workbooks.Open(filepath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

            Worksheet worksheet = (Worksheet)wb.Sheets["Лист1"];
            return worksheet;
            /*
            var a1 = worksheet.Cells[1,10];
            object rawValue = a1.Value;
            string formattedText = a1.Text;
            Console.WriteLine("rawValue={0} formattedText={1}", rawValue, formattedText);
            */
        }

    }
}
