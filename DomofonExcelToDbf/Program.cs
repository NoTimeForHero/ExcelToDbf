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

            var a = DateTime.ParseExact("Декабря 2017", "MMMM yyyy", CultureInfo.GetCultureInfo("ru-ru"));
            var b = DateTime.ParseExact("Декабрь 2017", "MMMM yyyy", CultureInfo.GetCultureInfo("ru-ru"));
            Console.WriteLine(a);
            Console.WriteLine(b);
            //createDBF();
            //readExcel();
            //readCOMExcel();
            //WriteResourceToFile("xConfig", "config.xml");
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

        private void readCOMExcel()
        {
            var app = new Microsoft.Office.Interop.Excel.Application();

            Workbook wb = app.Workbooks.Open(@"C:\test\example.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

            Worksheet worksheet = (Worksheet)wb.Sheets["Лист1"];
            var a1 = worksheet.Cells[1,10];
            object rawValue = a1.Value;
            string formattedText = a1.Text;
            Console.WriteLine("rawValue={0} formattedText={1}", rawValue, formattedText);
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
