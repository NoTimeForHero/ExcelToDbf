using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils;
using Microsoft.Office.Interop.Excel;
using NLog;
using SocialExplorer.IO.FastDBF;

namespace ExcelToDbf.Core.Services
{
    internal class DBFService
    {
        protected readonly ConfigProvider pvConfig;
        protected readonly ILogger logger;

        public DBFService(ILogger logger, ConfigProvider pvConfig)
        {
            this.logger = logger;
            this.pvConfig = pvConfig;
        }

        public Work Make(DocForm form, string outputFilename) => new Work(this, form, outputFilename);

        public class Work : IDisposable
        {
            private readonly DbfFile db;
            private readonly ILogger logger;
            private readonly DocForm form;
            private readonly List<DbfColumn> columns = new List<DbfColumn>();

            internal Work(DBFService owner, DocForm form, string outputFilename)
            {
                this.form = form;
                logger = owner.logger;
                var config = owner.pvConfig.Config;

                var encoding = Encoding.GetEncoding(config.System.OutputEncoding);

                db = new DbfFile(encoding);
                db.Open(outputFilename, FileMode.Create); // FileMode.Create = файл будет перезаписан если уже существует
                logger.Info($"Создаём DBF по пути: {outputFilename}");
                logger.Debug("и кодировкой: {encoding}");
            }

            public Work WriteHeaders()
            {
                foreach (var field in form.Fields)
                {
                    DbfColumn.DbfColumnType colType;
                    string length = field.Length;
                    var type = field.Type?.ToLower() ?? "string";
                    switch (type)
                    {
                        case "date":
                            colType = DbfColumn.DbfColumnType.Date;
                            length = length ?? "8";
                            break;
                        case "number":
                            colType = DbfColumn.DbfColumnType.Number;
                            length = length ?? "10,2";
                            break;
                        case "string":
                            colType = DbfColumn.DbfColumnType.Character;
                            break;
                        default:
                            throw new InvalidOperationException($"Unknown DBF field type: {field.Type}");
                    }
                    var lenParts = length.Split(',');
                    int nLen = int.Parse(lenParts[0]);
                    int nDec = lenParts.Length > 1 ? int.Parse(lenParts[1]) : 0;
                    var column = new DbfColumn(field.Name, colType, nLen, nDec);
                    db.Header.AddColumn(column);
                    columns.Add(column);
                    logger.Debug($"Записываем поле '{type}' типа '{type}' длиной {nLen},{nDec}");
                }

                db.WriteHeader();

                // config.System.OutputEncoding
                return this;
            }

            public void WriteRecord(Dictionary<string, object> input)
            {
                if (input == null) return;

                var record = new DbfRecord(db.Header);
                // record.AllowIntegerTruncate = true;
                record.AllowStringTurncate = true;

                for (int i=0; i<columns.Count; i++)
                {
                    var column = columns[i];
                    if (!input.TryGetValue(column.Name, out var rawValue)) continue;
                    switch (column.ColumnType)
                    {
                        case DbfColumn.DbfColumnType.Character:
                            record[i] = rawValue?.ToString();
                            break;
                        case DbfColumn.DbfColumnType.Date:
                            record[i] = DateHelper.ToDBF(rawValue?.ToString());
                            break;
                        case DbfColumn.DbfColumnType.Number:
                            record[i] = rawValue?.ToString().Replace(',', '.');
                            break;
                        default:
                            throw new NotImplementedException($"Column {column.ColumnType} not supported yet!");
                    }
                }

                db.Write(record, true);
            }

            public void Dispose()
            {
                db.Close();
            }
        }

    }
}
