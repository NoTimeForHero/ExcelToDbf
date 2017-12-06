using SocialExplorer.IO.FastDBF;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DomofonExcelToDbf.Sources.Core;
using DomofonExcelToDbf.Sources.Xml;

namespace DomofonExcelToDbf.Sources
{
    public class DBF
    {
        protected string path;
        protected DbfFile odbf;
        protected List<Xml_DbfField> dbfFields;
        protected int records;
        protected bool closed;
        protected bool headersWrited;

        public int Writed => records;

        public DBF(String path, List<Xml_DbfField> dbfFields, Encoding encoding = null)
        {
            this.dbfFields = dbfFields;
            this.path = path;

            if (encoding == null) encoding = Encoding.GetEncoding(866);

            odbf = new DbfFile(encoding);
            odbf.Open(path, FileMode.Create); // FileMode.Create = файл будет перезаписан если уже существует
            Logger.info($"Создаём DBF по пути: {path}");
            Logger.debug("и кодировкой: {encoding}");
        }

        public void writeHeader()
        {
            Logger.info($"Записываем в DBF {dbfFields.Count} полей:");
            Logger.info(string.Join(", ", dbfFields.Select(x => x.name).ToArray()));
            foreach (var field in dbfFields)
            {
                string name = field.name;
                string type = field.type;

                string attrlen = field.length;

                DbfColumn.DbfColumnType column = DbfColumn.DbfColumnType.Character;
                if (type == "string") column = DbfColumn.DbfColumnType.Character;
                if (type == "date") column = DbfColumn.DbfColumnType.Date;
                if (type == "numeric") column = DbfColumn.DbfColumnType.Number;

                if (attrlen != null)
                {
                    var length = attrlen.Split(',');
                    int nlen = Int32.Parse(length[0]);
                    int ndec = (length.Length > 1) ? Int32.Parse(length[1]) : 0;
                    odbf.Header.AddColumn(new DbfColumn(name, column, nlen, ndec));
                    Logger.debug($"Записываем поле '{name}' типа '{type}' длиной {nlen},{ndec}");
                }
                else
                {
                    odbf.Header.AddColumn(new DbfColumn(name, column));
                    Logger.debug($"Записываем поле '{name}' типа '{type}'");
                }
            }
            odbf.WriteHeader();
            headersWrited = true;
        }

        public void appendRecord(Dictionary<string, TVariable> variables)
        {
            if (!headersWrited) throw new Exception("Невозможно вставить запись в DBF раньше записи заголовков!");

            var orec = new DbfRecord(odbf.Header);
            //orec.AllowIntegerTruncate = true;
            orec.AllowStringTurncate = true;

            int fid = 0;
            foreach (var field in dbfFields)
            {
                string input = field.text;
                string type = field.type ?? "string";

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

                        object data = variables[repvar].value ?? "";

                        if (type == "string" || type == "numeric")
                        {
                            input = input.Replace(m.Value, data.ToString());
                            if (type == "numeric") input = input.Replace(',', '.');
                        }
                        else if (type == "date")
                        {
                            string format = field.format ?? "yyyy-MM-dd";
                            input = input.Replace(m.Value, ((DateTime)data).ToString(format));
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception($"Ошибка в переменной \"{input}\": {ex.Message}", ex);
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
            close();
            File.Delete(path);
        }

    }
}
