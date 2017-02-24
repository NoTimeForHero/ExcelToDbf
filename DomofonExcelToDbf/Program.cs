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

            readRecords(sheet, form);


            var a = DateTime.ParseExact("Декабря 2017", "MMMM yyyy", CultureInfo.GetCultureInfo("ru-ru"));
            //var b = DateTime.ParseExact("Декабрь 2017", "MMMM yyyy", CultureInfo.GetCultureInfo("ru-ru"));
            //Console.WriteLine(a);
            //Console.WriteLine(b);
            //createDBF();
            //readExcel();
            //readCOMExcel();
            //WriteResourceToFile("xConfig", "config.xml");
        }

        public void readRecords(Worksheet worksheet, XElement form)
        {
            Dictionary<string, object> variables = new Dictionary<string, object>();

            var staticvars = form.Element("Fields").Elements("Static");
            foreach (XElement staticvar in staticvars)
            {
                var x = Int32.Parse(staticvar.Attribute("X").Value);
                var y = Int32.Parse(staticvar.Attribute("Y").Value);

                var name = staticvar.Attribute("name").Value;
                String type = attrOrDefault(staticvar, "type", "string");

                Console.WriteLine(staticvar);

                var cell = worksheet.Cells[y, x].Value;
                if (cell == null)
                {
                    variables.Add(name, "");
                    continue;
                }

                if (type == "string")
                {
                    variables.Add(name, cell);
                }

                if (type == "date")
                {
                    var date = readDate(staticvar, cell);
                    variables.Add(name, date);
                }
                             
                if (cell != null) Console.WriteLine(cell);
                //inlineRead(worksheet, x, y, field);
            }

            Console.WriteLine("Составление записей завершено?");

        }

        public DateTime readDate(XElement staticvar, String cell)
        {
            var format = staticvar.Attribute("format").Value;
            var language = attrOrDefault(staticvar, "language", "ru-ru");
            DateTime date = DateTime.ParseExact(cell, format, CultureInfo.GetCultureInfo(language));

            // Если нам нужен последний день в месяце
            string lastday = attrOrDefault(staticvar, "lastday", "");
            if (lastday == "true") date = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            return date;
        }

        public string attrOrDefault(XElement element, String attr, String def) {
            XAttribute xattr = element.Attribute(attr);
            if (xattr == null) return def;
            return xattr.Value;
        }

        public object inlineRead(Worksheet worksheet, int x, int y, XElement field)
        {
            String type = field.Attribute("type").Value;
            if (type == null) type = "string";

            return null;
        }

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
