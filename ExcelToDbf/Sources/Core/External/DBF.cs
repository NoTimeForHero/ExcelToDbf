using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelToDbf.Sources.Core.Data;
using SocialExplorer.IO.FastDBF;

namespace ExcelToDbf.Sources.Core.External
{
    public class DBF
    {
        protected string path;
        protected DbfFile odbf;
        protected List<Xml_DbfField> dbfFields;
        protected int records;

        public int Writed => records;

        public DBF(String path, List<Xml_DbfField> dbfFields, Encoding encoding = null)
        {
            this.dbfFields = dbfFields;
            this.path = path;

            if (encoding == null) encoding = Encoding.GetEncoding(866);
            if (dbfFields == null) throw new ArgumentNullException(nameof(dbfFields));
            if (path == null ) throw new ArgumentNullException(nameof(path));

            odbf = new DbfFile(encoding);
            odbf.Open(path, FileMode.Create); // FileMode.Create = файл будет перезаписан если уже существует
            Logger.info($"Создаём DBF по пути: {path}");
            Logger.debug("и кодировкой: {encoding}");

            try
            {
                writeHeader();
            }
            catch (ArgumentNullException ex)
            {
                Logger.error("Исключение ArgumentNullException: " + ex.ParamName);
                delete();
                throw;
            }
        }

        protected void writeHeader()
        {
            Logger.info($"Записываем в DBF {dbfFields.Count} полей:");
            Logger.info(string.Join(", ", dbfFields.Select(x => x != null ? x.name : "[null!]").ToArray()));
            foreach (var field in dbfFields)
            {
                if (field == null) throw new ArgumentNullException(nameof(field));
                string name = field.name ?? throw new ArgumentNullException(nameof(field.name));
                string type = field.type ?? throw new ArgumentNullException(nameof(field.type));
                string attrlen = field.length ?? throw new ArgumentNullException(nameof(field.type));

                DbfColumn.DbfColumnType column = DbfColumn.DbfColumnType.Character;
                if (type == "string") column = DbfColumn.DbfColumnType.Character;
                if (type == "date") column = DbfColumn.DbfColumnType.Date;
                if (type == "numeric") column = DbfColumn.DbfColumnType.Number;

                var length = attrlen.Split(',');
                int nlen = Int32.Parse(length[0]);
                int ndec = (length.Length > 1) ? Int32.Parse(length[1]) : 0;
                odbf.Header.AddColumn(new DbfColumn(name, column, nlen, ndec));
                Logger.debug($"Записываем поле '{name}' типа '{type}' длиной {nlen},{ndec}");
            }
            odbf.WriteHeader();
        }

        public void appendRecord(Dictionary<string, TVariable> variables)
        {
            var orec = new DbfRecord(odbf.Header);
            //orec.AllowIntegerTruncate = true;
            orec.AllowStringTurncate = true;

            int fid = 0;
            foreach (var field in dbfFields)
            {
                string input = field.text;
                string type = field.type ?? "string";

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

                orec[fid] = input;
                fid++;
            }

            odbf.Write(orec, true);

            records++;
        }

        public void close()
        {
            odbf.Close();
        }

        public void delete()
        {
            close();
            File.Delete(path);
        }

    }
}
