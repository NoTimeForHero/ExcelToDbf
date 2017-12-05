using DomofonExcelToDbf.Sources;
using DomofonExcelToDbf.Sources.Xml;
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
using System.Xml.Serialization;
using DomofonExcelToDbf.Properties;
using DomofonExcelToDbf.Sources.Core;
using DomofonExcelToDbf.Sources.View;
using Application = System.Windows.Forms.Application;
using Point = System.Drawing.Point;

namespace DomofonExcelToDbf
{
    public class Program
    {

        [STAThread]
        static void Main(string[] args)
        {


            var exists = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().Location)).Count() > 1;
            if (exists)
            {
                MessageBox.Show("Программа уже запущена!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Распаковка DLL, которая не находится при упаковке через LibZ 
            File.WriteAllBytes("Microsoft.WindowsAPICodePack.dll", Resources.Microsoft_WindowsAPICodePack);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Program program = new Program();
            MainWindow window = new MainWindow(program);
            window.FormClosing += new FormClosingEventHandler(program.onFormMainClosing);
            Application.Run(window);
        }

        string confName;
        public XDocument xdoc;
        public Xml_Config config;
        public bool showStacktrace = false;
        Thread process;

        public Dictionary<string, string> formToFile = new Dictionary<string, string>();
        public List<string> outlog = new List<string>();
        public List<string> errlog = new List<string>();
        public HashSet<string> filesExcel = new HashSet<string>();
        public HashSet<string> filesDBF = new HashSet<string>();

        public void init()
        {
            confName = Path.ChangeExtension(AppDomain.CurrentDomain.FriendlyName, ".xml");

            if (!File.Exists(confName))
            {
                Console.WriteLine(@"Не найден конфигурационный файл!");
                Console.WriteLine(@"Распаковываем его из внутренних ресурсов...");
                WriteResourceToFile("xConfig", confName);
            }

            config = Xml_Config.Load(confName);
            xdoc = XDocument.Load(confName);

            if (!Directory.Exists("logs")) Directory.CreateDirectory("logs");
            Logger.instance = new Logger(config.log ? "logs\\" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log" : null);

            updateDirectory();

            Logger.instance.log("Версия программы: " + Resources.version);
        }

        public void updateDirectory()
        {
            Logger.instance.log("Директория чтения: {0}", config.inputDirectory);
            Logger.instance.log("Директория записи: {0}", config.outputDirectory);

            if (!Directory.Exists(config.inputDirectory)) config.inputDirectory = Directory.GetCurrentDirectory();
            if (!Directory.Exists(config.outputDirectory)) config.outputDirectory = Directory.GetCurrentDirectory();

            filesDBF.Clear();
            filesExcel.Clear();

            string[] fbyext = Directory.GetFiles(config.outputDirectory, "*.dbf", SearchOption.TopDirectoryOnly);
            filesDBF.UnionWith(fbyext);

            foreach (string extension in config.extensions)
            {
                fbyext = Directory.GetFiles(config.inputDirectory, extension, SearchOption.TopDirectoryOnly);
                fbyext = fbyext.Where(path => path != null
                        && !Path.GetFileName(path).Equals(confName) // А также наш конфигурационный файл %EXE_NAME%.xml
                        && !Path.GetFileName(path).StartsWith("~$")).ToArray(); // Игнорируем временные файлы Excel вида ~$Document.xls[x]
                filesExcel.UnionWith(fbyext);
            }
        }

        private void onFormMainClosing(object sender, FormClosingEventArgs e)
        {
            onCloseCheckProcess(e);

            xdoc.Root.Element("inputDirectory").Value = config.inputDirectory;
            xdoc.Root.Element("outputDirectory").Value = config.outputDirectory;
            xdoc.Save(confName);
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
            wstatus.Location = new Point(
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

        protected void delegate_action(object obj)
        {
            object [] data = (object[])obj;
                   
            StatusWindow window = (StatusWindow)data[0];
            MainWindow wmain = (MainWindow)data[1];
            HashSet<string> files = (HashSet<string>)data[2];
            window.setState(true, "Подготовка файлов", 0, files.Count);
            int idoc = 1;

            Excel excel = new Excel(config.save_memory);
            DBF dbf = null;

            var totalwatch = new Stopwatch();
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
                    window.updateState(true, $"Документ: {Path.GetFileName(finput)}", idoc);
                    idoc++;

                    excel.OpenWorksheet(finput);

                    var form = findCorrectForm(excel.worksheet, xdoc, config);

                    if (config.only_rules)
                    {
                        var formname = (form != null) ? form.Name : "null";
                        formToFile.Add(Path.GetFileName(finput), formname);
                        continue;
                    }

                    if (form == null)
                    {
                        Logger.instance.log("Не найдено подходящих форм для обработки документа work.xml!");
                        throw new NoNullAllowedException("Не найдено подходящих форм для обработки документа work.xml!");
                    }

                    string fileName = getOutputFilename(excel.worksheet, finput, config.outfile.simple, config.outfile.script);
                    string pathTemp = Path.GetTempFileName();
                    string pathOutput = Path.Combine(config.outputDirectory, fileName);

                    var total = excel.worksheet.UsedRange.Rows.Count - form.Fields.StartY;
                    window.setState(false, $"Обработано записей: {0}/{total}", 0, total);

                    dbf = new DBF(pathTemp,form.DBF);
                    dbf.writeHeader();

                    var stopwatch = new Stopwatch();

                    RegExCache cache = new RegExCache();

                    stopwatch.Start();
                    Work work = new Work(form, config.buffer_size);
                    work.IterateRecords(excel.worksheet, dbf.appendRecord, 
                        (int id) => window.updateState(false, $"Обработано записей: {id}/{total}", id)
                    );
                    stopwatch.Stop();

                    dbf.close();

                    Logger.instance.log("Времени потрачено на обработку данных: {0}", stopwatch.Elapsed);
                    Logger.instance.log("Обработано записей: {0} ", dbf.records);
                    outlog.Add(String.Format("{0} в {1} строк за {2}",Path.GetFileName(finput),dbf.records,stopwatch.Elapsed.ToString("hh\\:mm\\:ss\\.ff")));

                    int startY = form.Fields.StartY; 
                    Logger.instance.log("Начиная с {0} по {1}", startY, startY + dbf.records);

                    // Перемещение файла
                    if (File.Exists(pathOutput)) File.Delete(pathOutput);
                    File.Move(pathTemp, pathOutput);
                    Logger.instance.log($"Перемещение файла с {pathTemp} в {pathOutput}");

                    Logger.instance.log(String.Format("=============== Документ {0} успешно обработан! ===============", Path.GetFileName(finput)));
                }               
                /*
                catch (Exception ex)
                {
                    if (ex is ThreadAbortException)
                    { 
                        excel.close();
                        goto skip_error_msgbox;
                    }

                    errlog.Add(String.Format("Документ \"{0}\" был пропущен!",Path.GetFileName(finput)));

                    string stacktrace = (showStacktrace) ? ex.StackTrace : "";

                    var message = String.Format("Ошибка! Документ \"{0}\" будет пропущен!\n\n{1}\n\n{2}", Path.GetFileName(finput), ex.Message, stacktrace);
                    Logger.instance.log(message + "\n" + ex.StackTrace);
                    MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    skip_error_msgbox:;
                    Console.Error.WriteLine(ex);
                    deleteDbf = true;
                } 
                */
                finally
                {
                    Logger.instance.log("Закрытие COM Excel и DBF");
                    dbf?.close();
                    if (deleteDbf) dbf?.delete();
                }

            }
            totalwatch.Stop();

            // Не забываем завершить Excel
            excel.close();

            string crules = "";

            if (config.only_rules)
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
            crules += String.Format("\nВремени затрачено суммарно: {0}", totalwatch.Elapsed.ToString("hh\\:mm\\:ss\\.ff"));

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

        public string getOutputFilename(Worksheet worksheet, String inputFile, bool simple, string script = null)
        {
            if (simple) return Path.GetFileName(Path.ChangeExtension(inputFile, ".dbf"));

            JS.DelegateReadExcel readCell = (int x, int y) =>
            {
                try
                {
                    return worksheet.Cells[y, x].Value;
                }
                catch (Exception ex)
                {
                    Logger.instance.log($"Ошибка при чтении ячейки x={x},y={y}: {ex.Message}");
                    return null;
                }
            };

            JS js = new JS(readCell, Logger.instance.log);
            js.SetPath(inputFile);

            string outputFilename = js.Execute(script);
            if (!outputFilename.EndsWith(".dbf")) outputFilename += ".dbf";
            return outputFilename;
        }

        // <summary>
        // Ищет подходящую XML форму для документа или null если ни одна не подходит
        // </summary>
        public Xml_Form findCorrectForm(Worksheet worksheet, XDocument xdoc, Xml_Config config)
        {
            RegExCache regExCache = new RegExCache();

            foreach (Xml_Form form in config.Forms)
            {
                bool correct = true;
                Logger.instance.log($"\nПроверяем форму \"{form.Name}\"");
                Logger.instance.log("==========================================");

                foreach (Xml_Equal rule in form.Rules)
                {
                    bool useRegex = rule.regex_pattern != null;
                    bool validateRegex = rule.validate == "regex";

                    string cell = null;
                    string origcell = null;

                    try
                    {
                        cell = worksheet.Cells[rule.Y, rule.X].Value.ToString();
                    }
                    catch (Exception ex)
                    {
                        Logger.instance.log($"Произошла ошибка при чтении ячейки Y={rule.Y},X={rule.X}!");
                        Logger.instance.log($"Ожидалось: {rule.Text}");
                        Logger.instance.log("Ошибка: {0}", ex.Message);
                        correct = false;
                        break;
                    }

                    origcell = cell;
                    if (useRegex && !validateRegex)
                    {
                        cell = regExCache.MatchGroup(cell, rule.regex_pattern, rule.regex_group);
                    }

                    bool failed = false;
                    if (rule.Text != cell && !validateRegex) failed = true;
                    if (validateRegex && !regExCache.IsMatch(cell, rule.Text)) failed = true;

                    if (failed)
                    {
                        if (validateRegex || useRegex) Logger.instance.log("Провалена проверка по регулярному выражению!");
                        Logger.instance.log($"Проверка провалена (Y={rule.Y},X={rule.X})");
                        Logger.instance.log($"Ожидалось: {rule.Text}");
                        Logger.instance.log($"Найдено: {cell}");
                        if (useRegex)
                        {
                            Logger.instance.log($"Оригинальная ячейка: {origcell}");
                            Logger.instance.log($"Регулярное выражение: {rule.regex_pattern}");
                            Logger.instance.log($"Группа для поиска: {rule.regex_group}");
                        }
                        correct = false;
                        break;
                    }
                    Logger.instance.log($"Y={rule.Y},X={rule.X}: {rule.Text}{(validateRegex ? " is match" : "==")}{cell}");
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

    public class MyException : Exception
    {
        private readonly string myStackTrace;

        public MyException(string message, Exception exp) : base(message)
        {
            myStackTrace = exp.StackTrace;
        }

        public override string StackTrace => base.StackTrace + "\n" + myStackTrace;
    }

}