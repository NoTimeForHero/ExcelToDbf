using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DomofonExcelToDbf.Sources.Xml;

namespace DomofonExcelToDbf.Sources
{
    class Tools
    {

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
                }
                catch (Exception ex)
                {
                    Logger.instance.log(String.Format("Ошибка при чтении ячейки x={0},y={1}: {2}", x, y, ex.Message));
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
        public static Xml_Form findCorrectForm(Worksheet worksheet, XDocument xdoc, Xml_Config config)
        {
            var forms = xdoc.Root.Element("Forms").Elements("Form").ToList();
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
