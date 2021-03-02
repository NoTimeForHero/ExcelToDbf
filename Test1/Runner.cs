using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

// Работает только на NET Framework 3.5
// При перехода на 4.0 игнорирует delimiter в CSV файлах
namespace CSV_Converter
{
    public class Runner
    {
        private static string Exe => Assembly.GetExecutingAssembly().Location;

        public static void Main(string[] args)
        {
            if (args == null || args.Length < 2)
            {
                Console.WriteLine("Usage: %input_path%\\input.csv %output_path\\output.csv [delimiter=;]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string delimiter = (args.Length >= 3) ? args[2] : ";";
            if (args.Length >= 3) delimiter = args[2];
            SaveAs(inputPath, outputPath, delimiter);
        }

        public static string Open(string inputFilename, string delimiter)
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "ExcelToDbf");
            Directory.CreateDirectory(tempDir);

            var filename = Path.ChangeExtension(Path.GetRandomFileName(), ".xls");
            var outputFilename = Path.Combine(tempDir, filename);

            Run(inputFilename, outputFilename, delimiter).WaitForExit();
            if (!File.Exists(outputFilename)) return null;
            return outputFilename;
        }

        internal static Process Run(string inputPath, string outputPath, string delimiter=";")
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = Exe,
                Arguments = $"\"{inputPath}\" \"{outputPath}\" {delimiter}"
            };
            return Process.Start(startInfo);
        }

        internal static void SaveAs(string path, string output, string delimiter)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(
                    path,
                    Format: Excel.XlFileFormat.xlCSV,
                    Delimiter: delimiter
                );
                wb.SaveAs(output, Excel.XlFileFormat.xlExcel8);
                app.Visible = false;
            }
            finally
            {
                wb?.Close();
                app?.Quit();
            }
        }

    }
}
