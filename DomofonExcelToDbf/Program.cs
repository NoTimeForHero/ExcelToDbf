using Microsoft.Office.Interop.Excel;
using NickBuhro.Translit;
using SocialExplorer.IO.FastDBF;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
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


        public void appendRecord(Dictionary<string, TVariable> variables)
        {
            var orec = new DbfRecord(odbf.Header);
            //orec.AllowIntegerTruncate = true;
            orec.AllowStringTurncate = true;

            int fid = 0;
            foreach (XElement field in dbfields)
            {

                string input = field.Value;
                string name = field.Attribute("name").Value;
                string type = XmlHelper.attrOrDefault(field, "type", "string");

                try
                {

                    var matches = Regex.Matches(input, "\\$([0-9a-zA-Z]+)", RegexOptions.Compiled);
                    foreach (Match m in matches)
                    {
                        var repvar = m.Groups[1].Value;

                        if (!variables.ContainsKey(repvar)) // чтобы в финальном файле не оказалось строк вида $VARIABLE
                        {
                            input = input.Replace(m.Value, "");
                            continue;
                        }

                        object data = variables[repvar].value;
                        if (data == null) data = "";

                        if (type == "string" || type == "numeric")
                        {
                            input = input.Replace(m.Value, data.ToString());
                            if (type == "numeric") input = input.Replace(',', '.');
                        }
                        else if (type == "date")
                        {
                            string format = XmlHelper.attrOrDefault(field, "format", "yyyy-MM-dd");
                            input = input.Replace(m.Value, ((DateTime)data).ToString(format));
                        }
                    }

                }                
                catch (Exception ex)
                {
                    throw new Exception(String.Format("Ошибка в переменной \"{0}\": {1}", input, ex.Message), ex);
                }              

                orec[fid] = input;
                fid++;
            }

            odbf.Write(orec, true);

            records++;
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

    class XmlHelper
    {
        public static string attrOrDefault(XElement element, String attr, String def)
        {
            XAttribute xattr = element.Attribute(attr);
            if (xattr == null) return def;
            return xattr.Value;
        }

        public static String attr(XElement element, String attr)
        {
            XAttribute xattr = element.Attribute(attr);
            if (xattr == null) return null;
            return xattr.Value;
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

        string confName;
        public XDocument xdoc;
        public bool onlyRules;
        public bool showStacktrace;
        public bool saveMemory;
        public String dirInput;
        public String dirOutput;
        public String status;
        public String labelTitle;

        Thread process = null;

        public Dictionary<string, string> formToFile = new Dictionary<string, string>();
        public List<string> outlog = new List<string>();
        public List<string> errlog = new List<string>();
        public HashSet<string> filesExcel = new HashSet<string>();
        public HashSet<string> filesDBF = new HashSet<string>();
        public int record_buffer;

        public void init()
        {
            confName = Path.ChangeExtension(System.AppDomain.CurrentDomain.FriendlyName, ".xml");

            if (!File.Exists(confName))
            {
                Console.WriteLine("Не найден конфигурационный файл!");
                Console.WriteLine("Распаковываем его из внутренних ресурсов...");
                Tools.WriteResourceToFile("xConfig", confName);
            }

            xdoc = XDocument.Load(confName);

            var log = xdoc.Root.Element("log");
            bool is_log = (log != null && log.Value == "true");
            string log_file = !is_log ? null : "logs\\" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log";
            if (!Directory.Exists("logs")) Directory.CreateDirectory("logs");
            Logger.instance = new Logger(log_file);

            var status = xdoc.Root.Element("status");
            this.status = (status != null) ? status.Value : "";

            var xStackTrace = xdoc.Root.Element("show_stacktrace");
            this.showStacktrace = (xStackTrace != null && xStackTrace.Value == "true");

            dirInput = xdoc.Root.Element("inputDirectory").Value; 
            dirOutput = xdoc.Root.Element("outputDirectory").Value;
            labelTitle = xdoc.Root.Element("title") != null ? xdoc.Root.Element("title").Value : "";

            record_buffer = xdoc.Root.Element("buffer_size") != null ? Int32.Parse(xdoc.Root.Element("buffer_size").Value) : 200;

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
            }
        }

        private void onFormMainClosing(object sender, FormClosingEventArgs e)
        {
            onCloseCheckProcess(e);

            xdoc.Root.Element("inputDirectory").Value = dirInput;
            xdoc.Root.Element("outputDirectory").Value = dirOutput;
            xdoc.Save(confName);

            #if !DEBUG
                File.Delete("Microsoft.WindowsAPICodePack.dll");
            #endif
        }

        private void onCloseCheckProcess(FormClosingEventArgs e)
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

        public void action(MainWindow wmain, HashSet<string> files)
        {
            if (process != null && process.IsAlive)
            {
                MessageBox.Show("Процесс конвертирования уже запущен!\nДождись его завершения, если вы хотите начать новый.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            StatusWindow wstatus = new StatusWindow();
            wstatus.FormClosing += new FormClosingEventHandler(delegate(object sender, FormClosingEventArgs e)
            {
                if (e.CloseReason != CloseReason.UserClosing) return;
                if (wstatus.codeClose) return;
                e.Cancel = DialogResult.No == MessageBox.Show("Вы действительно хотите прервать обработку файлов?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (!e.Cancel)
                {
                    process.Abort();
                    wstatus.Hide();
                    MessageBox.Show(wmain, "Документы не были обработаны: процесс был прерван пользователем!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
            wstatus.Location = new System.Drawing.Point(
                wmain.Location.X + ((wmain.Width - wstatus.Width) / 2),
                wmain.Location.Y + ((wmain.Height - wstatus.Height) / 2)
            );
            wstatus.Show(wmain);
            // Альтернативный вариант:
            //wstatus.StartPosition = FormStartPosition.CenterParent;
            //wstatus.ShowDialog(wmain);

            object data = new object[3] { wstatus, wmain, files };

            outlog.Clear();
            errlog.Clear();
            formToFile.Clear();

            process = new Thread(delegate_action);
            process.Start(data);
        }

        private void Wstatus_FormClosing(object sender, FormClosingEventArgs e)
        {
            throw new NotImplementedException();
        }

        protected void delegate_action(object obj)
        {
            object [] data = (object[])obj;
                   
            StatusWindow window = (StatusWindow)data[0];
            MainWindow wmain = (MainWindow)data[1];
            HashSet<string> files = (HashSet<string>)data[2];
            window.setState(true, "Подготовка файлов", 0, files.Count);
            int idoc = 1;

            Excel excel = new Excel(saveMemory);
            DBF dbf = null;

            var totalwatch = new System.Diagnostics.Stopwatch();
            totalwatch.Start();
            foreach (string fname in files)
            {

                // COM Excel требуется полный путь до файла
                string finput = Path.GetFullPath(fname);

                bool deleteDbf = false;

                try
                {
                    Logger.instance.log("\n");
                    Logger.instance.log("==============================================================");
                    Logger.instance.log("Загружаем Excel документ: {0}", Path.GetFileName(finput));
                    Logger.instance.log("==============================================================");
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

                    string fileName = Tools.getOutputFilename(excel.worksheet, xdoc, dirInput, finput);
                    string pathTemp = Path.GetTempFileName();
                    string pathOutput = Path.Combine(dirOutput, fileName);

                    var total = excel.worksheet.UsedRange.Rows.Count - Tools.startY(form);
                    window.setState(false, String.Format("Обработано записей: {0}/{1}", 0, total), 0, total);

                    dbf = new DBF(pathTemp);
                    dbf.writeHeader(form);

                    var stopwatch = new System.Diagnostics.Stopwatch();

                    RegExCache cache = new RegExCache();

                    stopwatch.Start();
                    Work work = new Work(xdoc,form, record_buffer);
                    work.IterateRecords(excel.worksheet, dbf.appendRecord, 
                        (int id) => window.updateState(false, String.Format("Обработано записей: {0}/{1}", id, total), id)
                    );
                    stopwatch.Stop();

                    dbf.close();

                    Logger.instance.log("Времени потрачено на обработку данных: {0}", stopwatch.Elapsed);
                    Logger.instance.log("Обработано записей: {0} ", dbf.records);
                    outlog.Add(String.Format("{0} в {1} строк за {2}",Path.GetFileNameWithoutExtension(finput),dbf.records,stopwatch.Elapsed.ToString("hh\\:mm\\:ss\\.ff")));

                    int startY = Tools.startY(form);    
                    Logger.instance.log("Начиная с {0} по {1}", startY, startY + dbf.records);

                    // Перемещение файла
                    if (File.Exists(pathOutput)) File.Delete(pathOutput);
                    File.Move(pathTemp, pathOutput);
                    Logger.instance.log(string.Format("Перемещение файла с {0} в {1}", pathTemp, pathOutput));

                    Logger.instance.log(string.Format("=============== Документ {0} успешно обработан! ===============", Path.GetFileName(finput)));
                }               
                catch (Exception ex)
                {
                    if (ex is ThreadAbortException)
                    { 
                        excel.close();
                        goto skip_error_msgbox;
                    }

                    errlog.Add(String.Format("Документ \"{0}\" был пропущен!",Path.GetFileNameWithoutExtension(finput)));

                    string stacktrace = (showStacktrace) ? ex.StackTrace : "";

                    var message = String.Format("Ошибка! Документ \"{0}\" будет пропущен!\n\n{1}\n\n{2}", Path.GetFileNameWithoutExtension(finput), ex.Message, stacktrace);
                    Logger.instance.log(message + "\n" + ex.StackTrace);
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
                    string line = String.Format("Для \"{0}\" форма \"{1}\"", tup.Key, tup.Value);
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

                var xmlWarning = xdoc.Root.Element("warning");          
                string warnFormat = (xmlWarning == null) ? "{0}" : xmlWarning.Value;
                warnFormat = warnFormat.Replace("\\n", "\n");
                crules += String.Format(warnFormat,String.Join("\n", errlog));
            }

            updateDirectory();
            wmain.BeginInvoke((MethodInvoker)wmain.fillElementsData);

            window.mayClose();
            MessageBox.Show(crules, "Отчёт о времени обработки", MessageBoxButtons.OK, icon);
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

        public void log()
        {
            log("");
        }

        public void log(object data)
        {
            Console.WriteLine(data.ToString());
            if (!console)
            {
                writer.WriteLine(data.ToString());
                writer.Flush();
            }
        }

        public void log(string data, object arg0, object arg1=null, object arg2=null, object arg3=null)
        {
            Console.WriteLine(data, arg0, arg1, arg2, arg3);
            if (!console)
            {
                writer.WriteLine(data, arg0, arg1, arg2, arg3);
                writer.Flush();
            }
        }

    }

    class RegExCache
    {
        protected Dictionary<String,Regex> regexes = new Dictionary<String,Regex>();

        protected Regex Prepare(String strregex)
        {
            if (!regexes.ContainsKey(strregex)) regexes.Add(strregex, new Regex(strregex, RegexOptions.IgnoreCase | RegexOptions.Compiled));
            return regexes[strregex];
        }

        public String Replace(String input, String strregex, String replacement="$1")
        {
            Regex regex = Prepare(strregex);
            return regex.Replace(input, replacement);
        }

        public bool IsMatch(String input, String strregex)
        {
            Regex regex = Prepare(strregex);
            return regex.Match(input).Success;
        }

        public String MatchGroup(String input, String strregex, int group=1)
        {
            Regex regex = Prepare(strregex);
            Match match = regex.Match(input);
            if (!match.Success) return "";
            if (match.Groups.Count - 1 < group) return "";
            return match.Groups[group].Value;
        }

        public static String MatchGroup(String input, Regex regex, int group = 1)
        {
            Match match = regex.Match(input);
            if (!match.Success) return "";
            if (match.Groups.Count - 1 < group) return "";
            return match.Groups[group].Value;
        }

    }

    public class Work
    {
        public Dictionary<string,TVariable> staticVars = new Dictionary<string, TVariable>();
        public Dictionary<string, TVariable> dynamicVars  = new Dictionary<string, TVariable>();
        public HashSet<TCondition> conditions = new HashSet<TCondition>();

        protected int startY;
        protected int endX;
        protected int buffer;
        protected int total = 0;
        protected XElement form;
        protected TVariable exception_var;

        public Dictionary<string, TVariable> stepScope = new Dictionary<string, TVariable>();
        
        public Work(XDocument xdocument, XElement form, int buffer)
        {
            InitVariables(form);
            startY = Tools.startY(form);
            endX = Tools.endX(form);
            this.buffer = buffer;
            this.form = form;
        }

        public void IterateRecords(Worksheet worksheet, Action<Dictionary<string, TVariable>> callback, Action<int> guiCallback = null)
        {
            total = 0;
            try
            {
                __IterateRecords(worksheet, callback, guiCallback);
            } catch (Exception ex)
            {
                string message = string.Format("Ошибка на строке {0}, ячейке {1} в переменной {2}:\n{3}", startY + total, exception_var.x, exception_var.name, ex.Message);
                throw new MyException(message, ex);
            }
            FinalChecks();
        }

        protected void __IterateRecords(Worksheet worksheet, Action<Dictionary<string, TVariable>> callback, Action<int> guiCallback = null)
        {
            int begin = startY;
            int end = startY + buffer;

            var maxY = worksheet.UsedRange.Rows.Count;

            Stopwatch watch;
            bool EOF = false;

            watch = Stopwatch.StartNew();
            stepScope.Clear();
            foreach (var var in staticVars.Values)
            {
                exception_var = var;
                var.Set(worksheet.Cells[var.y, var.x].Value);
                stepScope.Add(var.name, var);
            }
            watch.Stop();
            Logger.instance.log("Заполнение массива локальных переменных: " + watch.ElapsedMilliseconds);

            Stopwatch watchTotal = Stopwatch.StartNew();
            while (!EOF)
            {
                var range_start = worksheet.Cells[begin, 1];
                var range_end = worksheet.Cells[end, endX];
                var range = worksheet.Range[range_start, range_end];
                object[,] tmp = range.Value;

                watch = Stopwatch.StartNew();
                for (int i = 1; i <= buffer; i++)
                {
                    bool skipRecord = false;
                    bool stopLoop = false;

                    foreach (TCondition cond in conditions)
                    {
                        if (cond.mustBe.Equals(tmp[i, cond.x]) || (cond.mustBe == "" && tmp[i, cond.x] == null))
                        {
                            foreach (TAction item in cond.onTrue)
                            {
                                if (item is TInterrupt tinter)
                                {
                                    if (tinter.action == TInterrupt.Action.SKIP_RECORD)
                                    {
                                        Console.WriteLine(String.Format("Пропуск записи по условию: значение в ячейке x={0} равно {1}", cond.x, cond.mustBe));
                                        skipRecord = true;
                                    }
                                    if (tinter.action == TInterrupt.Action.STOP_LOOP)
                                    {
                                        Console.WriteLine(String.Format("Выход из цикла по условию: значение в ячейке x={0} равно {1}", cond.x, cond.mustBe));
                                        stopLoop = true;
                                    }
                                    continue;
                                }
                                if (item is TVariable var)
                                {
                                    exception_var = var;
                                    var.Set(tmp[i, var.x]);
                                    stepScope[var.name] = var;
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            foreach (TAction item in cond.onFalse)
                            {
                                if (item is TInterrupt tinter)
                                {
                                    if (tinter.action == TInterrupt.Action.SKIP_RECORD)
                                    {
                                        Console.WriteLine(String.Format("Пропуск записи по условию: значение в ячейке x={0} равно {1}", cond.x, cond.mustBe));
                                        skipRecord = true;
                                    }
                                    if (tinter.action == TInterrupt.Action.STOP_LOOP)
                                    {
                                        Console.WriteLine(String.Format("Выход из цикла по условию: значение в ячейке x={0} равно {1}", cond.x, cond.mustBe));
                                        stopLoop = true;
                                    }
                                    continue;
                                }
                                if (item is TVariable var)
                                {
                                    exception_var = var;
                                    var.Set(tmp[i, var.x]);
                                    stepScope[var.name] = var;
                                    continue;
                                }
                            }
                        }
                    }

                    total++;

                    if (total > maxY - startY)
                    {
                        Logger.instance.log("Попытка выйти за пределы документа, выход из цикла");
                        EOF = true;
                        break;
                    }

                    if (stopLoop)
                    {
                        Logger.instance.log("Выход из цикла по условию");
                        EOF = true;
                        break;                        
                    }

                    if (skipRecord) continue;

                    foreach (var var in dynamicVars.Values)
                    {
                        exception_var = var;
                        var.Set(tmp[i, var.x]);
                        stepScope[var.name] = var;
                    }

                    callback(stepScope);
                    guiCallback?.Invoke(total);
                }
                watch.Stop();
                Logger.instance.log(String.Format("Сегмент в {0} элементов (с {1} по {2}) обработан за {3} мс", buffer, begin, end, watch.ElapsedMilliseconds));

                begin += buffer;
                end += buffer;
            }
            watchTotal.Stop();
            Logger.instance.log("Total time: " + watchTotal.ElapsedMilliseconds);
            Logger.instance.log("Rows iterated: " + total);
            Logger.instance.log("Buffer size:" + buffer);
        }

        protected void FinalChecks()
        {
            int num = 1;

            if (form.Element("Validate") == null) return;
            foreach (XElement validate in form.Element("Validate").Elements())
            {
                stepScope.TryGetValue(validate.Attribute("var1").Value, out TVariable var1);
                stepScope.TryGetValue(validate.Attribute("var2").Value, out TVariable var2);

                string value1 = (var1 == null || var1.value == null) ? "[неизвестно]" : var1.value.ToString();
                string value2 = (var2 == null || var2.value == null) ? "[неизвестно]" : var2.value.ToString();

                var elemMsg = validate.Element("Message");
                string message = "";

                if (elemMsg == null)
                {
                    message = string.Format("Финальная проверка №{0} провалена!", num);
                } else
                {
                    message = string.Format(elemMsg.Value, value1, value2, num);
                    message = message.Replace("\\n", "\n");
                }

                if (var1 == null || var2 == null || var1.value == null || var2.value == null) throw new Exception(message);

                Logger.instance.log(string.Format(
                    "Проверка номер {0} : {1}({2}) сравнивается с {3}({4})",
                    num, var1 != null ? var1.name : "null", value1, var2 != null ? var2.name : "null", value2));

                bool isEqual = false;
                if (validate.Element("Math") is XElement math && math.Attribute("type").Value == "numeric")
                {
                    int count = Int32.Parse(math.Attribute("count").Value);
                    float prec = Single.Parse(math.Attribute("precision").Value);

                    float allowed_precision = (prec / count) * total;
                    float var1fl = Convert.ToSingle(var1.value);
                    float var2fl = Convert.ToSingle(var2.value);

                    Logger.instance.log("var1 = " + var1fl.ToString("G9"));
                    Logger.instance.log("var2 = " + var2fl.ToString("G9"));

                    if (var1fl == var2fl) isEqual = true;
                    else
                    {
                        float diff = Math.Abs(Math.Abs(var1fl) - Math.Abs(var2fl));
                        isEqual = diff < allowed_precision;
                        message += "\n" + string.Format(math.Value, allowed_precision, diff);
                        Logger.instance.log(string.Format(math.Value, allowed_precision, diff));
                    }
                }
                else isEqual = var1.value.Equals(var2.value);

                if (!isEqual) throw new Exception(message);

                num++;
            }
        }

        protected void InitVariables(XElement form)
        {
            foreach (XElement xelem in form.Element("Fields").Elements())
            {
                if (xelem.Name == "Static") AddVar(staticVars,getVar(xelem,false));
                if (xelem.Name == "Dynamic") AddVar(dynamicVars,getVar(xelem,true));
                if (xelem.Name == "IF") conditions.Add(ScanCondition(xelem));
            }
        }

        protected TCondition ScanCondition(XElement xml)
        {
            TCondition condition = new TCondition();
            condition.x = Int32.Parse(xml.Attribute("X").Value);
            condition.mustBe = xml.Attribute("VALUE").Value;

            foreach (XElement elem in xml.Element("THEN").Elements()) {
                TAction action = null;
                if (elem.Name == "SKIP_RECORD")
                    action = new TInterrupt(TInterrupt.Action.SKIP_RECORD);
                if (elem.Name == "STOP_LOOP")
                    action = new TInterrupt(TInterrupt.Action.STOP_LOOP);
                if (elem.Name == "Dynamic")
                    action = getVar(elem,true);
                if (action != null) condition.onTrue.Add(action);
            }

            if (xml.Element("ELSE") != null)
            {
                foreach (XElement elem in xml.Element("ELSE").Elements())
                {
                    TAction action = null;
                    if (elem.Name == "SKIP_RECORD")
                        action = new TInterrupt(TInterrupt.Action.SKIP_RECORD);
                    if (elem.Name == "STOP_LOOP")
                        action = new TInterrupt(TInterrupt.Action.STOP_LOOP);
                    if (elem.Name == "Dynamic")
                        action = getVar(elem, true);
                    if (action != null) condition.onFalse.Add(action);
                }
            }
            return condition;
        }

        protected void AddVar(IDictionary<string, TVariable>  dictionary, TVariable variable)
        {
            dictionary.Add(variable.name, variable);
        }

        protected TVariable getVar(XElement xml, bool dynamic)
        {
            var name = xml.Attribute("name").Value;
            var ctype = (xml.Attribute("type") != null) ? xml.Attribute("type").Value : "string";

            TVariable.Type type = TVariable.getByString(ctype);
            TVariable variable;
            switch (type)
            {
                case TVariable.Type.ENumeric:
                    variable = new TNumeric(name);
                    break;
                case TVariable.Type.EDate:
                    variable = new TDate(name);
                    break;
                default:
                    variable = new TVariable(name);
                    break;
            }

            variable.x = Int32.Parse(xml.Attribute("X").Value);
            if (!dynamic) variable.y = Int32.Parse(xml.Attribute("Y").Value);
            variable.dynamic = dynamic;
            variable.type = type;

            if (variable is TNumeric tnumeric)
            {
                if (xml.Attribute("function") != null)
                    tnumeric.function = TNumeric.getFuncByString(xml.Attribute("function").Value);
            }

            if (variable is TDate tdate)
            {
                if (xml.Attribute("lastday") != null)
                    tdate.lastday = Boolean.Parse(xml.Attribute("lastday").Value);
                if (xml.Attribute("language") != null)
                    tdate.language = xml.Attribute("language").Value;
                if (xml.Attribute("format") != null)
                    tdate.format = xml.Attribute("format").Value;
            }

            var regex_pattern = xml.Attribute("regex_pattern");
            if (regex_pattern != null)
            {
                variable.use_regex = true;
                variable.regex_pattern = new Regex(regex_pattern.Value, RegexOptions.Compiled);
                variable.regex_group = xml.Attribute("regex_group") != null ? Int32.Parse(xml.Attribute("regex_group").Value) : 1;
            }
            return variable;
        }
    }

    /// <summary>
    /// Универсальный класс, от которого наследуются всё возможные операции
    /// Сюда входят: условия, переменные, прерывания цикла обработки
    /// </summary>
    public abstract class TAction {}

    public class TInterrupt : TAction
    {
        public Action action;

        public TInterrupt(Action action)
        {
            this.action = action;
        }

        public enum Action
        {
            SKIP_RECORD,
            STOP_LOOP
        }
    }

    public class TCondition : TAction
    {
        public int x;
        public string mustBe;

        public List<TAction> onTrue = new List<TAction>();
        public List<TAction> onFalse = new List<TAction>();
    }

    public class TVariable : TAction
    {
        public enum Type : byte
        {
            EUnknown,
            EString,
            ENumeric,
            EDate
        }

        public readonly string name;
        public Type type;

        public bool dynamic;
        public int x;
        public int y;

        public object value;

        public bool use_regex = false;
        public Regex regex_pattern;
        public int regex_group;

        public TVariable(string name)
        {
            this.name = name;
        }

        public void Set(object val)
        {
            string str = (val == null) ? "" : val.ToString();
            if (use_regex)
                str = RegExCache.MatchGroup(str, regex_pattern, regex_group);

            if (false) ;
            else if (this is TDate tdate) tdate.Set(str);
            else if (this is TNumeric tnum) tnum.Set(str);
            else this.value = str;
        }

        public static Type getByString(string str)
        {
            if (str == "string") return Type.EString;
            if (str == "date") return Type.EDate;
            if (str == "numeric") return Type.ENumeric;
            return Type.EUnknown;
        }

        public override bool Equals(object obj)
        {
            var item = obj as TVariable; 
            if (item == null) return false;
            return this.name == item.name;
        }

        public override int GetHashCode()
        {
            return name.GetHashCode();
        }
    }

    public class TNumeric : TVariable
    {
        public Func function = Func.NONE;

        public enum Func : byte
        {
            NONE,
            SUM
        }

        public static Func getFuncByString(string str)
        {
            if (str == "SUM") return Func.SUM;
            return Func.NONE;
        }

        public TNumeric(string name) : base(name) { }

        public new void Set(object obj)
        {
            if ("".Equals(obj)) obj = "0"; // Иначе Convert.ToSingle упадёт с ошибкой
            float value = Convert.ToSingle(obj);
            switch (function)
            {
                case Func.SUM:
                    this.value = Convert.ToSingle(this.value) + value;
                    break;
                default:
                    this.value = value;
                    break;
            }
        }
    }

    public class TDate : TVariable
    {
        public bool lastday = false;
        public string format = "dd.MM.yyy";
        public string language = "ru-ru";

        public TDate(string name) : base(name) {}

        public new void Set(object val)
        {
            DateTime date = DateTime.ParseExact(val as string, format, CultureInfo.GetCultureInfo(language));
            if (lastday) date = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            this.value = date;
        }
    }
    
    class Tools {         

        public static string getOutputFilename(Worksheet worksheet, XDocument xdoc, String inputDirectory, String inputFile)
        {
            XElement outfile = xdoc.Root.Element("outfile");

            bool simple = outfile.Element("simple").Value == "true";
            if (simple) return Path.GetFileName(Path.ChangeExtension(inputFile, ".dbf"));

            string script = outfile.Element("script").Value;

            JS.DelegateReadExcel readCell = (int x, int y) =>
            {
                try
                {
                    return worksheet.Cells[y, x].Value;
                } catch (Exception ex)
                {
                    Logger.instance.log(String.Format("Ошибка при чтении ячейки x={0},y={1}: {2}",x,y,ex.Message));
                    return null;
                }
            };

            JS js = new JS(readCell, Logger.instance.log);
            js.SetPath(inputFile);

            string outputFilename = js.Execute(script);
            if (!outputFilename.EndsWith(".dbf")) outputFilename += ".dbf";
            return outputFilename;
        }

        public static Int32 startY(XElement form)
        {
            var val = form.Element("Fields").Element("StartY");
            if (val == null) throw new ArgumentNullException("Required tag <StartY> in <Fields> section is null!");
            return Int32.Parse(val.Value);
        }

        public static Int32 endX(XElement form)
        {
            var val = form.Element("Fields").Element("EndX");
            if (val == null) throw new ArgumentNullException("Required tag <EndX> in <Fields> section is null!");
            return Int32.Parse(val.Value);
        }

        // <summary>
        // Ищет подходящую XML форму для документа или null если ни одна не подходит
        // </summary>
        public static XElement findCorrectForm(Worksheet worksheet, XDocument xdoc)
        {
            var forms = xdoc.Root.Element("Forms").Elements("Form").ToList();
            RegExCache regExCache = new RegExCache();

            foreach (XElement form in forms)
            {
                bool correct = true;
                String name = form.Element("Name").Value;
                Logger.instance.log(String.Format("\nПроверяем форму \"{0}\"",name));
                Logger.instance.log("==========================================");

                var equals = form.Element("Rules").Elements("Equal");
                foreach (XElement equal in equals)
                {
                    var x = Int32.Parse(equal.Attribute("X").Value);
                    var y = Int32.Parse(equal.Attribute("Y").Value);
                    var mustbe = equal.Value;

                    bool useRegex = equal.Attribute("regex_pattern") != null;
                    string regex_pattern = useRegex ? equal.Attribute("regex_pattern").Value.ToString() : "";
                    int regex_group = equal.Attribute("regex_group") != null ? Int32.Parse(equal.Attribute("regex_group").Value) : 1;
                    bool validateRegex = equal.Attribute("validate") != null && equal.Attribute("validate").Value.ToString() == "regex";

                    string cell = null;
                    string origcell = null;

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

                    origcell = cell;
                    if (useRegex && !validateRegex) {
                        cell = regExCache.MatchGroup(cell, regex_pattern, regex_group);
                    }                    

                    bool failed = false;
                    if (mustbe != cell && !validateRegex) failed = true;
                    if (validateRegex && !regExCache.IsMatch(cell, mustbe)) failed = true;

                    if (failed)
                    {
                        if (validateRegex || useRegex) Logger.instance.log("Провалена проверка по регулярному выражению!");
                        Logger.instance.log(String.Format("Проверка провалена (Y={0},X={1})",y,x));
                        Logger.instance.log(String.Format("Ожидалось: {0}", mustbe));
                        Logger.instance.log(String.Format("Найдено: {0}", cell));
                        if (useRegex)
                        {
                            Logger.instance.log(String.Format("Оригинальная ячейка: {0}", origcell));
                            Logger.instance.log(String.Format("Регулярное выражение: {0}", regex_pattern));
                            Logger.instance.log(String.Format("Группа для поиска: {0}", regex_group));
                        }
                        correct = false;
                        break;
                    }
                        Logger.instance.log(String.Format("Y={0},X={1}: {2}{4}{3}",y,x,mustbe,cell,(validateRegex? " is match" : "==")));
                    }
                    if (correct) return form;
            }
            return null;
        }

        // <summary>
        // Метод считывает внутренний ресурс и записывает его в файл, возвращая статус существования ресурса
        // </summary>
        // <param name="resourceName">Имя внутренного ресурса</param>
        // <param name="fileName">Имя внутренного ресурса</param>
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

    class JS
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
            engine.SetValue("dir", new System.Action(() => FuncThrowException("Ошибка 1754: Невозможно выполнить функцию dir(...), так как не установлена директория до конечного файла через JS->SetPath(...)!")));
        }

        public string Execute(string script)
        {
            return engine.Execute(script).GetCompletionValue().ToObject().ToString();
        }

        public void SetPath(string fullPath)
        {
            DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(fullPath));
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

    /// <summary>
    /// Класс для удобного разбития пути на сегменты
    /// Например C:\One\Two\Three превратятся в массив: "C:\", "One", "Two", "Three"
    /// Любой элемент из которого можно получить через метод GetLevel(index)
    /// </summary>
    public class PathHelper
    {

        public List<string> components;

        public int Count
        {
            get
            {
                return components.Count;
            }
        }

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
            if (path == null) throw new ArgumentNullException("path");
            var ret = new List<string>();
            if (path.Parent != null) ret.AddRange(Split(path.Parent));
            ret.Add(path.Name);
            return ret;
        }
    }

    public class MyException : Exception
    {
        private string myStackTrace;

        public MyException(string message, Exception exp) : base(message)
        {
            this.myStackTrace = exp.StackTrace;
        }

        public override string StackTrace
        {
            get
            {
                return base.StackTrace + "\n" + myStackTrace;
            }
        }
    }
}