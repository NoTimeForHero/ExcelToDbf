using Microsoft.Office.Interop.Excel;
using SocialExplorer.IO.FastDBF;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
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

                    var cell = worksheet.Cells[y, x].Value.ToString();

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

        private void readExcel()
        {
            Console.WriteLine("Где первая строка? ");


            string connectionString = GetConnectionString();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data   
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        Console.WriteLine("Начинаем читать файл...");
                        Console.WriteLine();
                        while (reader.Read())
                        {
                            var count = reader.FieldCount;
                            for (int i = 0; i < count; i++)
                            {
                                var x = reader.GetValue(i);
                                Console.Write(x);
                                Console.Write('|');
                            }
                            Console.WriteLine();
                        }
                    }
                }
                cmd = null;
                conn.Close();
            }
        }

        private string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            //props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            //props["Extended Properties"] = "Excel 12.0 XML";
            //props["Extended Properties"] = "\"Excel 12.0 Xml; HDR=No\"";
            //props["Data Source"] = @"C:\Users\user\Documents\visual studio 2015\Projects\DomofonExcelToDbf\DomofonExcelToDbf\bin\Debug\example.xls";

            //props["Driver"] = "{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb, *.xml)};";
            //props["DBQ"] = @"C:\Users\user\Documents\visual studio 2015\Projects\DomofonExcelToDbf\DomofonExcelToDbf\bin\Debug\example.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

    }
}
