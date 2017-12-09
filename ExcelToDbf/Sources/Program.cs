using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using ExcelToDbf.Properties;
using ExcelToDbf.Sources.Core;
using ExcelToDbf.Sources.Core.Data.Xml;
using ExcelToDbf.Sources.Core.External;
using ExcelToDbf.Sources.View;
using Application = System.Windows.Forms.Application;
using Point = System.Drawing.Point;
using System.Runtime.InteropServices;

namespace ExcelToDbf.Sources
{
    [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
    public class Program
    {
        public static readonly bool DEBUG = Debugger.IsAttached;

        [STAThread]
        private static void Main()
        {
            bool exists = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().Location)).Length > 1;
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
            window.FormClosing += program.onFormMainClosing;
            Application.Run(window);
        }

        readonly string confName;
        public XDocument xdoc;
        public Xml_Config config;
        public bool showStacktrace = false;
        private Thread process;

        public Dictionary<string, string> formToFile = new Dictionary<string, string>();
        public List<string> outlog = new List<string>();
        public List<string> errlog = new List<string>();
        public HashSet<string> filesExcel = new HashSet<string>();
        public HashSet<string> filesDBF = new HashSet<string>();

        public Program()
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

            if (config.log) Logger.SetFile("logs\\" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log");
            Logger.ParseLevel(config.LogLevel);

            updateDirectory();

            Logger.info("Версия программы: " + Resources.version);
            Logger.info("Уровень логирования: " + Logger.Level);
        }

        /// <summary>
        /// Обновляет список файлов для установленных в конфиге директорий
        /// </summary>
        public void updateDirectory()
        {
            Logger.debug("Директория чтения: " + config.inputDirectory);
            Logger.debug("Директория записи: " + config.outputDirectory);

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

        /// <summary>
        /// Запускает процесс конвертирования файлов в отдельном потоке
        /// Вызывается по кнопке "Конвертировать" на форме
        /// </summary>
        /// <param name="wmain">Окно, которое вызывает процесс конвертации</param>
        /// <param name="selectedfiles">Список файлов для конвертирования (с учётом выбора пользователя)</param>
        public void action(MainWindow wmain, HashSet<string> selectedfiles)
        {
            if (selectedfiles.Count > 0)
            {
                DialogResult ask = MessageBox.Show("Вы действительно хотите конвертировать только выбранные файлы?", "Вопрос",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (ask == DialogResult.No) selectedfiles = filesExcel;
                if (ask == DialogResult.Cancel) return;
            }
            else selectedfiles = filesExcel;

            if (selectedfiles.Count == 0 && filesExcel.Count == 0)
            {
                MessageBox.Show($"В директории нет Excel файлов для обработки!\nВыберите другую директорию!\n\n{config.inputDirectory}",
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (process != null && process.IsAlive)
            {
                MessageBox.Show("Процесс конвертирования уже запущен!\nДождись его завершения, если вы хотите начать новый.", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            StatusWindow wstatus = new StatusWindow();
            wstatus.FormClosing += delegate(object sender, FormClosingEventArgs e)
            {
                if (e.CloseReason != CloseReason.UserClosing) return;
                if (wstatus.codeClose) return;
                e.Cancel = DialogResult.No == MessageBox.Show("Вы действительно хотите прервать обработку файлов?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (!e.Cancel)
                {
                    process.Abort();
                    wstatus.mayClose();
                    MessageBox.Show(wmain, "Документы не были обработаны: процесс был прерван пользователем!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            wstatus.Location = new Point(
                wmain.Location.X + ((wmain.Width - wstatus.Width) / 2),
                wmain.Location.Y + ((wmain.Height - wstatus.Height) / 2)
            );
            wstatus.Show(wmain);
            // Альтернативный вариант:
            //wstatus.StartPosition = FormStartPosition.CenterParent;
            //wstatus.ShowDialog(wmain);

            object data = new object[] { wstatus, wmain, selectedfiles.ToList() };

            outlog.Clear();
            errlog.Clear();
            formToFile.Clear();

            process = new Thread(delegate_action);
            process.Start(data);
        }

        /// <summary>
        /// Основной поток обработки
        /// </summary>
        /// <param name="obj">Форма процесса обработки, главная форма, список файлов для обработки</param>
        protected void delegate_action(object obj)
        {
            object [] data = (object[])obj;

            StatusWindow window = (StatusWindow)data[0];
            MainWindow wmain = (MainWindow)data[1];
            List<string> files = (List<string>)data[2];
            window.setState(true, "Подготовка файлов", 0, files.Count);

            if (config.only_rules)
            {
                CheckRules(files, window);
                return;
            }

            bool reopenExcel = true;
            Excel excel = null;
            DBF dbf = null;

            Stopwatch totalwatch = Stopwatch.StartNew();
            for (int idoc=0;idoc<files.Count;idoc++)
            {
                string pathFull = files[idoc];
                string filename = Path.GetFileName(pathFull);
                string pathTemp = Path.GetTempFileName();

                bool deleteDbf = false;

                if (reopenExcel)
                {
                    excel = new Excel();
                    reopenExcel = false;
                }

                try
                {

                    Logger.info("");
                    Logger.debug("");
                    Logger.debug("==============================================================");
                    Logger.info($"======= Загружаем Excel документ: {filename} ======");
                    Logger.debug("==============================================================");
                    window.updateState(true, $"Документ: {filename}", idoc);

                    excel.OpenWorksheet(pathFull);

                    var form = findCorrectForm(excel.worksheet, config);
                    if (form == null)
                    {
                        Logger.warn("Не найдено подходящих форм для обработки документа: " + filename);
                        throw new NoNullAllowedException($"Не найдено подходящих форм для обработки документа '{filename}'!");
                    }

                    string fileName = getOutputFilename(excel.worksheet, pathFull, config.outfile.simple, config.outfile.script);
                    string pathOutput = Path.Combine(config.outputDirectory, fileName);

                    dbf = new DBF(pathTemp,form.DBF);

                    var total = excel.worksheet.UsedRange.Rows.Count - form.Fields.StartY;
                    window.setState(false, $"Обработано записей: {0}/{total}", 0, total);

                    Work work = new Work(form, config.buffer_size);
                    TimeSpan elapsed = work.IterateRecords(excel.worksheet, dbf.appendRecord,
                        id => window.updateState(false, $"Обработано записей: {id}/{total}", id)
                    );

                    dbf.close();

                    // Перемещение файла
                    if (File.Exists(pathOutput)) File.Delete(pathOutput);
                    File.Move(pathTemp, pathOutput);
                    Logger.debug($"Перемещение файла с {pathTemp} в {pathOutput}");

                    outlog.Add($"{filename} в {dbf.Writed} строк за {elapsed:hh\\:mm\\:ss\\.ff}");

                    Logger.info("Времени потрачено на обработку данных: " + elapsed);
                    Logger.info("Обработано записей: " + dbf.Writed);
                    Logger.debug($"Начиная с {form.Fields.StartY} по {form.Fields.StartY + dbf.Writed}");
                    Logger.info($"=============== Документ {Path.GetFileName(pathFull)} успешно обработан! ===============");
                }
                catch (Exception ex) when (!DEBUG)
                {
                    if (ex is COMException)
                    {
                        Logger.error("Excel вероятнее всего крашанулся, он будет перезапущен для следующего документа в очереди!");
                        reopenExcel = true;
                    }

                    if (ex is ThreadAbortException || ex.InnerException is ThreadAbortException)
                    {
                        Logger.warn($"Пользователь вышел во время процесса конвертации документа '{filename}'!");
                        goto skip_error_msgbox;
                    }

                    string stacktrace = (showStacktrace) ? $"\n\n{ex.StackTrace}" : "";
                    string message = $"Ошибка! Документ \"{filename}\" будет пропущен!\n\n{ex.Message}{stacktrace}";
                    errlog.Add($"Документ \"{filename}\" был пропущен!");

                    Logger.error($"Документ {filename} был пропущен из-за ошибки:\n{ex.Message}\n\n{ex.StackTrace}");
                    MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    skip_error_msgbox:;
                    deleteDbf = true;
                }
                finally
                {
                    Logger.debug("Закрытие COM Excel и DBF");
                    dbf?.close();
                    if (deleteDbf) dbf?.delete();
                }

            }
            totalwatch.Stop();

            // Не забываем завершить Excel
            excel.close();

            string crules = "Время обработки документов:\n";
            crules += String.Join("\n", outlog) + "\n";
            crules += String.Format("\nВремени затрачено суммарно: " + totalwatch.Elapsed.ToString("hh\\:mm\\:ss\\.ff"));
            Logger.info(crules);

            var icon = errlog.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning;
            if (errlog.Count > 0) {
                string warnFormat = config.warning ?? "{0}";
                warnFormat = warnFormat.Replace("\\n", "\n");
                crules += "\n\n" + String.Format(warnFormat,string.Join("\n", errlog));
            }

            updateDirectory();
            wmain.BeginInvoke((MethodInvoker)wmain.fillElementsData);
            window.mayClose();
            MessageBox.Show(crules, "Отчёт о времени обработки", MessageBoxButtons.OK, icon);
        }

        protected void CheckRules(IList<string> files, StatusWindow window)
        {
            if (!config.only_rules) return;
            string message = "";
            Excel excel = new Excel();

            window.setState(true, "", 0, files.Count);

            for (int id=0;id<files.Count;id++)
            {
                string fname = files[id];
                string docname = Path.GetFileName(fname) ?? fname;

                window.updateState(true, "Читаем документ " + docname, id+1);
                excel.OpenWorksheet(fname);

                var form = findCorrectForm(excel.worksheet, config);
                var formname = form?.Name ?? "[NULL]";
                string line = $"Для документа '{docname}' выбрана форма '{formname}'!";
                message += "\n" + line;
                Logger.info(line);
            }

            excel.close();
            window.mayClose();
            MessageBox.Show(message, "Отчёт о формах", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Метод, получающий имя выходного файла на основании JS скрипта из XML конфига
        /// </summary>
        /// <param name="worksheet">Excel документ для чтения ячеек в JS скрипте</param>
        /// <param name="inputFile">Полный путь к исходному Exel файлу</param>
        /// <param name="simple">Необходимо ли вообще использовать JS</param>
        /// <param name="script">Содержимое JS скрипта в виде текста</param>
        /// <returns></returns>
        public string getOutputFilename(Worksheet worksheet, String inputFile, bool simple, string script = null)
        {
            if (simple) return Path.GetFileName(Path.ChangeExtension(inputFile, ".dbf"));

            JS.DelegateReadExcel readCell = (x, y) =>
            {
                try
                {
                    return worksheet.Cells[y, x].Value;
                }
                catch (Exception ex)
                {
                    Logger.warn($"Ошибка при чтении ячейки x={x},y={y}: {ex.Message}");
                    return null;
                }
            };

            JS js = new JS(readCell, Logger.info);
            js.SetPath(inputFile);

            string outputFilename = js.Execute(script);
            if (!outputFilename.EndsWith(".dbf")) outputFilename += ".dbf";
            return outputFilename;
        }

        // <summary>
        // Ищет подходящую XML форму для документа или null если ни одна не подходит
        // </summary>
        [SuppressMessage("ReSharper", "ReplaceWithSingleAssignment.False")]
        public Xml_Form findCorrectForm(Worksheet worksheet, Xml_Config pConfig)
        {
            RegExCache regExCache = new RegExCache();

            foreach (Xml_Form form in pConfig.Forms)
            {
                bool correct = true;
                Logger.info("");
                Logger.info($"Проверяем форму \"{form.Name}\"");
                Logger.debug("==========================================");

                int index = 1;
                foreach (Xml_Equal rule in form.Rules)
                {
                    bool useRegex = rule.regex_pattern != null;
                    bool validateRegex = rule.validate == "regex";

                    string cell;

                    try
                    {
                        cell = worksheet.Cells[rule.Y, rule.X].Value.ToString();
                    }
                    catch (Exception ex)
                    {
                        Logger.debug($"Произошла ошибка при чтении ячейки Y={rule.Y},X={rule.X}!");
                        Logger.debug($"Ожидалось: {rule.Text}");
                        Logger.debug("Ошибка: " + ex.Message);
                        Logger.info($"Форма не подходит по условию №{index}");
                        correct = false;
                        break;
                    }

                    string origcell = cell;
                    if (useRegex && !validateRegex)
                    {
                        cell = regExCache.MatchGroup(cell, rule.regex_pattern, rule.regex_group);
                    }

                    bool failed = false;
                    if (rule.Text != cell && !validateRegex) failed = true;
                    if (validateRegex && !regExCache.IsMatch(cell, rule.Text)) failed = true;

                    if (failed)
                    {
                        if (validateRegex || useRegex) Logger.debug("Провалена проверка по регулярному выражению!");
                        Logger.debug($"Проверка провалена (Y={rule.Y},X={rule.X})");
                        Logger.debug($"Ожидалось: {rule.Text}");
                        Logger.debug($"Найдено: {cell}");
                        if (useRegex)
                        {
                            Logger.debug($"Оригинальная ячейка: {origcell}");
                            Logger.debug($"Регулярное выражение: {rule.regex_pattern}");
                            Logger.debug($"Группа для поиска: {rule.regex_group}");
                        }
                        Logger.info($"Форма не подходит по условию №{index}");
                        correct = false;
                        break;
                    }
                    Logger.debug($"Y={rule.Y},X={rule.X}: {rule.Text}{(validateRegex ? " is match" : "==")}{cell}");
                    index++;
                }
                if (correct)
                {
                    Logger.info($"Форма '{form.Name}' подходит для документа!");
                    return form;
                }
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
            using (var resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
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