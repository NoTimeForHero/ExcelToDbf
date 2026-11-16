using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

// Работает только на NET Framework 3.5
// При перехода на 4.0 игнорирует delimiter в CSV файлах
namespace CSV_Converter
{
    public class Runner
    {
        private static string Exe => Assembly.GetExecutingAssembly().Location;

        [STAThread]
        public static int Main(string[] args)
        {
            if (args == null || args.Length < 2)
            {
                Console.WriteLine("Usage: %input_path%\\input.csv %output_path\\output.csv [delimiter=;]");
                return 2;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string delimiter = (args.Length >= 3) ? args[2] : ";";
            if (args.Length >= 3) delimiter = args[2];
            try
            {
                SaveAs(inputPath, outputPath, delimiter, args.Length >= 4 ? args[3] : null);
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                return 1;
            }
        }

        public static ConvResult Open(string inputFilename, string delimiter)
            => Open(inputFilename, delimiter, null);

        public static ConvResult Open(string inputFilename, string delimiter, string ownershipDirectory)
        {
            var tempDir = ownershipDirectory ?? Path.Combine(Path.GetTempPath(), "ExcelToDbf");
            Directory.CreateDirectory(tempDir);

            var filename = Path.ChangeExtension(Path.GetRandomFileName(), ".xls");
            var outputFilename = Path.Combine(tempDir, filename);

            var result = new ConvResult {Filename = outputFilename};
            using (var process = Run(inputFilename, outputFilename, delimiter, ownershipDirectory))
            {
                var error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                if (process.ExitCode != 0) throw new IOException("Ошибка чтения CSV: " + error);
                result.Success = File.Exists(outputFilename);
            }
            return result;
        }

        internal static Process Run(string inputPath, string outputPath, string delimiter=";", string ownershipDirectory = null)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = Exe,
                Arguments = $"\"{inputPath}\" \"{outputPath}\" {delimiter}" + (ownershipDirectory == null ? "" : " \"" + ownershipDirectory + "\""),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true
            };
            var process = Process.Start(startInfo);
            RecordOwner(ownershipDirectory, process);
            return process;
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);

        private static void RecordOwner(string directory, Process process)
        {
            if (directory == null) return;
            File.WriteAllText(Path.Combine(directory, "owner-" + process.Id + ".txt"), process.Id + "|" +
                process.StartTime.ToUniversalTime().Ticks + "|" + process.ProcessName);
        }

        internal static void SaveAs(string path, string output, string delimiter, string ownershipDirectory)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Workbooks books = null;
            try
            {
                app = new Excel.Application();
                GetWindowThreadProcessId(new IntPtr(app.Hwnd), out var pid);
                using (var process = Process.GetProcessById((int)pid)) RecordOwner(ownershipDirectory, process);
                app.Visible = false;
                app.DisplayAlerts = false;
                app.AskToUpdateLinks = false;
                app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                books = app.Workbooks;
                wb = books.Open(
                    path,
                    Format: Excel.XlFileFormat.xlCSV,
                    Delimiter: delimiter, ReadOnly: true, UpdateLinks: 0
                );
                wb.SaveAs(output, Excel.XlFileFormat.xlExcel8);
                app.Visible = false;
            }
            finally
            {
                try { wb?.Close(false); }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                    if (books != null) Marshal.ReleaseComObject(books);
                    try { app?.Quit(); }
                    finally { if (app != null) Marshal.ReleaseComObject(app); }
                }
            }
        }

        public class ConvResult
        {
            public string Filename;
            public bool Success;
        }

    }
}
