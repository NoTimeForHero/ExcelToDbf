using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

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
        public static XElement findCorrectForm(Worksheet worksheet, XDocument xdoc)
        {
            var forms = xdoc.Root.Element("Forms").Elements("Form").ToList();
            RegExCache regExCache = new RegExCache();

            foreach (XElement form in forms)
            {
                bool correct = true;
                String name = form.Element("Name").Value;
                Logger.instance.log(String.Format("\nПроверяем форму \"{0}\"", name));
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
                    }
                    catch (Exception ex)
                    {
                        Logger.instance.log(String.Format("Произошла ошибка при чтении ячейки Y={0},X={1}!", y, x));
                        Logger.instance.log(String.Format("Ожидалось: {0}", mustbe));
                        Logger.instance.log("Ошибка: {0}", ex.Message);
                        correct = false;
                        break;
                    }

                    origcell = cell;
                    if (useRegex && !validateRegex)
                    {
                        cell = regExCache.MatchGroup(cell, regex_pattern, regex_group);
                    }

                    bool failed = false;
                    if (mustbe != cell && !validateRegex) failed = true;
                    if (validateRegex && !regExCache.IsMatch(cell, mustbe)) failed = true;

                    if (failed)
                    {
                        if (validateRegex || useRegex) Logger.instance.log("Провалена проверка по регулярному выражению!");
                        Logger.instance.log(String.Format("Проверка провалена (Y={0},X={1})", y, x));
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
                    Logger.instance.log(String.Format("Y={0},X={1}: {2}{4}{3}", y, x, mustbe, cell, (validateRegex ? " is match" : "==")));
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
