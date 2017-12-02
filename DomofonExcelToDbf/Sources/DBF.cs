using SocialExplorer.IO.FastDBF;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DomofonExcelToDbf.Sources
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
                }
                else
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
}
