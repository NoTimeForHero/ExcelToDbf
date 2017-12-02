﻿using DomofonExcelToDbf.Sources;
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
}