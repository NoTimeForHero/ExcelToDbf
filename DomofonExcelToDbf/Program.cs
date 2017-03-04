using Microsoft.Office.Interop.Excel;
using NickBuhro.Translit;
using SocialExplorer.IO.FastDBF;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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
            Logger.instance.log("Создаём DBF с именем {0} и\nкодировкой: {1}", path, encoding);
        }

        // Эту функцию нельзя вызвать за пределами данного класса
        public void writeHeader(XElement form)
        {
            dbfields = form.Element("DBF").Elements("field");
            Logger.instance.log("Записываем в DBF {0} полей", dbfields.Count());
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
                    Logger.instance.log("Записываем поле '{0}' типа '{1}' длиной {2},{3}", name, type, nlen, ndec);
                } else
                {
                    odbf.Header.AddColumn(new DbfColumn(name, column));
                    Logger.instance.log("Записываем поле '{0}' типа '{1}'", name, type);
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
            //if (i < 20) foreach (var x in variables) Logger.instance.log(x.Key + "=" + x.Value);

            records++;
            if (records % 100 > 0) return;
            Logger.instance.log("Записей обработано: {0}", records);
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
                Logger.instance.log("Выбранный Excel не содержит ни одного листа!");
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

    public class Program
    {
        [STAThread]
        static void Main(string[] args)
        {

            var exists = System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1;
            if (exists)
            {
                MessageBox.Show("Программа уже запущена!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Распаковка DLL, которая не находится при упаковке через LibZ 
            File.WriteAllBytes("Microsoft.WindowsAPICodePack.dll", DomofonExcelToDbf.Properties.Resources.Microsoft_WindowsAPICodePack);

            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            Program program = new Program();
            MainWindow window = new MainWindow(program);
            window.FormClosing += new FormClosingEventHandler(program.onFormMainClosing);
            System.Windows.Forms.Application.Run(window);
        }

        public XDocument xdoc;
        public bool onlyRules;
        public bool saveMemory;
        public String dirInput;
        public String dirOutput;
        public String status;

        Thread process = null;

        public Dictionary<string, string> formToFile = new Dictionary<string, string>();
        public List<string> outlog = new List<string>();
        public List<string> errlog = new List<string>();
        public HashSet<string> filesExcel = new HashSet<string>();
        public HashSet<string> filesDBF = new HashSet<string>();

        public void init()
        {
            String confName = Path.ChangeExtension(System.AppDomain.CurrentDomain.FriendlyName, ".xml");

            if (!File.Exists(confName))
            {
                Console.WriteLine("Не найден конфигурационный файл!");
                Console.WriteLine("Распаковываем его из внутренних ресурсов...");
                Tools.WriteResourceToFile("xConfig", confName);
            }

            xdoc = XDocument.Load(confName);

            var log = xdoc.Root.Element("log");
            string log_file = (log != null && log.Value == "true") ? Path.ChangeExtension(confName, ".log") : null;
            Logger.instance = new Logger(log_file);

            var status = xdoc.Root.Element("status");
            this.status = (status != null) ? status.Value : "";

            dirInput = xdoc.Root.Element("inputDirectory").Value; 
            dirOutput = xdoc.Root.Element("outputDirectory").Value;

            onlyRules = xdoc.Root.Element("only_rules").Value == "true";
            saveMemory = xdoc.Root.Element("save_memory").Value == "true"; // экономить память, если включено то будет использоваться один инстанс COM Excel с переключением Worksheet
            updateDirectory();
        }

        public void updateDirectory()
        {
            Logger.instance.log("Директория чтения: {0}", dirInput);
            Logger.instance.log("Директория записи: {0}", dirOutput);

            if (!Directory.Exists(dirInput)) dirInput = Directory.GetCurrentDirectory();
            if (!Directory.Exists(dirOutput)) dirOutput = Directory.GetCurrentDirectory();

            filesDBF.Clear();
            filesExcel.Clear();

            string[] fbyext = Directory.GetFiles(dirOutput, "*.dbf", SearchOption.TopDirectoryOnly);
            filesDBF.UnionWith(fbyext);

            foreach (var extension in xdoc.Root.Element("extensions").Elements("ext"))
            {
                fbyext = Directory.GetFiles(dirInput, extension.Value, SearchOption.TopDirectoryOnly);
                fbyext = fbyext.Where(path => !Path.GetFileName(path).StartsWith("~$")).ToArray(); // Игнорируем временные файлы Excel вида ~$Document.xls[x]
                filesExcel.UnionWith(fbyext);
                Logger.instance.log("Файлов найдено {1} по маске {0}", extension.Value, fbyext.Length);
            }
        }

        private void onFormMainClosing(object sender, FormClosingEventArgs e)
        {
            if (process == null) return;
            DialogResult abort = DialogResult.None;

            if (process.IsAlive)
            {
                abort = MessageBox.Show("Вы действительно хотите выйти?\nПроцесс конвертирования будет прерван.", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            }

            if (abort == DialogResult.No)
            {
                    e.Cancel = true;
                    return;
            }

            process.Abort();
        }

        public void action(MainWindow wmain)
        {
            if (process != null && process.IsAlive)
            {
                MessageBox.Show("Процесс конвертирования уже запущен!\nДождись его завершения, если вы хотите начать новый.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            StatusWindow wstatus = new StatusWindow();
            wstatus.Show();

            object forms = new object[2] { wstatus, wmain };

            process = new Thread(delegate_action);
            process.Start(forms);
        }

        protected void delegate_action(object obj)
        {
            object []forms = (object[])obj;
                   
            StatusWindow window = (StatusWindow)forms[0];
            MainWindow wmain = (MainWindow)forms[1];
            window.setState(true, "Подготовка файлов", 0, filesExcel.Count);
            int idoc = 1;

            Excel excel = new Excel(saveMemory);
            DBF dbf = null;

            var totalwatch = new System.Diagnostics.Stopwatch();
            totalwatch.Start();
            foreach (string fname in filesExcel)
            {

                // COM Excel требуется полный путь до файла
                string finput = Path.GetFullPath(fname);

                bool deleteDbf = false;

                try
                {
                    Logger.instance.log("Загружаем Excel документ: {0}", Path.GetFileName(finput));
                    window.updateState(true, String.Format("Документ: {0}", Path.GetFileName(finput)), idoc);
                    idoc++;

                    excel.OpenWorksheet(finput);

                    var form = Tools.findCorrectForm(excel.worksheet, xdoc);

                    if (onlyRules)
                    {
                        var formname = (form == null) ? "null" : form.Element("Name").Value;
                        formToFile.Add(Path.GetFileName(finput), formname);
                        continue;
                    }

                    if (form == null)
                    {
                        Logger.instance.log("Не найдено подходящих форм для обработки документа work.xml!");
                        throw new NoNullAllowedException("Не найдено подходящих форм для обработки документа work.xml!");
                    }

                    string foutput = Path.Combine(dirOutput, Tools.getOutputFilename(excel.worksheet, xdoc, dirInput, finput));

                    var total = excel.worksheet.UsedRange.Rows.Count - Tools.startY(form);
                    window.setState(false, String.Format("Обработано записей: {0}/{1}", 0, total), 0, total);

                    dbf = new DBF(foutput);
                    dbf.writeHeader(form);

                    var stopwatch = new System.Diagnostics.Stopwatch();

                    stopwatch.Start();
                    Tools.eachRecord(excel.worksheet, form, dbf.appendRecord, delegate(int id) { window.updateState(false, String.Format("Обработано записей: {0}/{1}", id, total), id); } );
                    stopwatch.Stop();

                    Logger.instance.log("Времени потрачено на обработку данных: {0}", stopwatch.Elapsed);
                    Logger.instance.log("Обработано записей: {0} ", dbf.records);
                    outlog.Add(String.Format("{0} в {1} строк за {2}",Path.GetFileNameWithoutExtension(finput),dbf.records,stopwatch.Elapsed.ToString("hh\\:mm\\:ss\\.ff")));

                    int startY = Tools.startY(form);    
                    Logger.instance.log("Начиная с {0} по {1}", startY, startY + dbf.records);
                }
                catch (Exception ex)
                {
                    if (ex is ThreadAbortException)
                    { 
                        excel.close();
                        goto skip_error_msgbox;
                    }

                    errlog.Add(String.Format("Документ \"{0}\" был пропущен!",Path.GetFileNameWithoutExtension(finput)));

                    var message = String.Format("Ошибка! Документ \"{0}\" будет пропущен!\n\n{1}", Path.GetFileNameWithoutExtension(finput), ex.Message);
                    Logger.instance.log(message);
                    MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    skip_error_msgbox:;
                    Console.Error.WriteLine(ex);
                    deleteDbf = true;
                }
                finally
                {
                    Logger.instance.log("Закрытие COM Excel и DBF");
                    if (dbf != null) dbf.close();
                    if (dbf != null && deleteDbf) dbf.delete();
                }

            }
            totalwatch.Stop();

            // Не забываем завершить Excel
            excel.close();

            string crules = "";

            if (onlyRules)
            {
                for (int i = 0; i < 3; i++) Logger.instance.log();
                foreach (var tup in formToFile)
                {
                    string line = String.Format("Для файла {0} выбрана форма {1}", tup.Key, tup.Value);
                    Logger.instance.log(line);
                    crules += line + "\n";
                }
                MessageBox.Show(crules, "Отчёт о формах", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            crules = "Время обработки документов:\n";
           
            var coutlog = String.Join("\n", outlog) + "\n";
            crules += coutlog;
            Logger.instance.log(coutlog);

            Logger.instance.log("Времени затрачено суммарно: {0}", totalwatch.Elapsed);
            crules += String.Format("Времени затрачено суммарно: {0}", totalwatch.Elapsed.ToString("hh\\:mm\\:ss\\.ff"));

            var icon = MessageBoxIcon.Information;

            if (errlog.Count > 0) {
                icon = MessageBoxIcon.Warning;
                crules += "\n\n";
                crules += String.Join("\n", errlog);
            }

            updateDirectory();
            wmain.BeginInvoke((MethodInvoker)wmain.fillElementsData);

            window.mayClose();
            MessageBox.Show(crules, "Отчёт о времени", MessageBoxButtons.OK, icon);
        }
    }

    class Logger
    {
        bool console = false;
        StreamWriter writer;

        public static Logger instance;

        public Logger(string file=null)
        {
            this.console = (file == null);
            if (file != null)
            {
                writer = new StreamWriter(file, false);
                writer.AutoFlush = true;
            }
        }

        public void log(string data="", object arg0=null, object arg1=null, object arg2=null, object arg3=null)
        {
            if (console) Console.WriteLine(data, arg0, arg1, arg2, arg3);
            else
            {
                writer.WriteLine(data, arg0, arg1, arg2, arg3);
                writer.Flush();
            }
        }

    }
    
    class Tools {         

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

        public static void eachRecord(Worksheet worksheet, XElement form, Action<Dictionary<string,object>> callback, Action<int> guiCallback = null)
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
                int id = y - minY;

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
                            Logger.instance.log("Ошибка в переменной {0} на Y={1},X={2}", name, y, x);
                            throw;
                        }
                    }

                    if (section.Element("SKIP_RECORD") != null)
                    {
                        Logger.instance.log("Пропускаем строку Y={0}", y);
                        goto skip_record;
                    }
                    if (section.Element("STOP_LOOP") != null)
                    {
                        Logger.instance.log("Выходим из цикла на Y={0} по условию X[{1}]={2}", y, cond.x, cond.value);
                        goto skip_loop;
                    }
                }

                callback(variables);
                if (id % 100 == 0) guiCallback?.Invoke(id);
                skip_record:;
            }
            skip_loop:;

            Logger.instance.log("Составление записей завершено?");

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
                Logger.instance.log(String.Format("Проверяем форму \"{0}\"",name));

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
                        Logger.instance.log(String.Format("Произошла ошибка при чтении ячейки Y={0},X={1}!", y, x));
                        Logger.instance.log(String.Format("Ожидалось: {0}", mustbe));
                        Logger.instance.log("Ошибка: {0}",ex.Message);
                        correct = false;
                        break;
                    }

                    if (mustbe != cell)
                        {
                            Logger.instance.log(String.Format("Проверка провалена (Y={0},X={1})",y,x));
                            Logger.instance.log(String.Format("Ожидалось: {0}", mustbe));
                            Logger.instance.log(String.Format("Найдено: {0}", cell));
                            correct = false;
                            break;
                        }
                        Logger.instance.log(String.Format("Y={0},X={1}:  {2}=={3}",y,x,mustbe,cell));
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