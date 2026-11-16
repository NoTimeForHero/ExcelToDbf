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
using ExcelToDbf.Processing;
using System.Globalization;

namespace ExcelToDbf.Core.Services
{
    public class DBFService
    {
        protected readonly ConfigProvider pvConfig;
        protected readonly ILogger logger;
        public Action<DataWarning> OnWarning { get; set; }
        public bool Strict { get; set; }

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
            private readonly DBFService owner;
            private readonly Encoding encoding;
            public int RecordsWritten { get; private set; }

            internal Work(DBFService owner, DocForm form, string outputFilename)
            {
                this.form = form;
                this.owner = owner;
                logger = owner.logger;
                var config = owner.pvConfig.Config;

                encoding = Encoding.GetEncoding(config.System.OutputEncoding);

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

            public void WriteRecord(Dictionary<string, object> input, int rowIndex)
            {
                if (input == null) return;

                var record = new DbfRecord(db.Header);
                // record.AllowIntegerTruncate = true;
                record.AllowStringTurncate = true;

                for (int i=0; i<columns.Count; i++)
                {
                    var column = columns[i];
                    if (!input.TryGetValue(column.Name, out var rawValue)) continue;
                    try
                    {
                    switch (column.ColumnType)
                    {
                        case DbfColumn.DbfColumnType.Character:
                            var text = rawValue?.ToString();
                            if (text != null)
                            {
                                var strictEncoding = (Encoding)encoding.Clone();
                                strictEncoding.EncoderFallback = EncoderFallback.ExceptionFallback;
                                try { strictEncoding.GetBytes(text); }
                                catch (EncoderFallbackException)
                                { Warn("ENCODING_LOSS", "Символы не представимы в выходной кодировке", rowIndex, column.Name, text); }
                                if (encoding.GetByteCount(text) > column.Length)
                                    Warn("STRING_TRUNCATED", "Строка превышает длину DBF-поля и будет обрезана", rowIndex, column.Name, text);
                            }
                            record[i] = text;
                            break;
                        case DbfColumn.DbfColumnType.Date:
                            var dateText = rawValue?.ToString();
                            var date = DbfHelper.ToDate(dateText);
                            if (!string.IsNullOrEmpty(dateText) && (date == null ||
                                !DateTime.TryParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _)))
                            {
                                Warn("INVALID_DATE", "Дата не распознана и будет записана пустой", rowIndex, column.Name, dateText);
                                date = null;
                            }
                            record[i] = date;
                            break;
                        case DbfColumn.DbfColumnType.Number:
                            var number = rawValue?.ToString().Replace(',', '.') ?? "";
                            if (!DbfHelper.IsNumber(number))
                                throw new ProcessingException("INVALID_NUMBER", "dbf-write", $"На строке {rowIndex} колонка {column.Name} ({number}) не является числом")
                                { Field = column.Name, SourceRow = rowIndex, Actual = number };
                            var spec = (form.Fields[i].Length ?? "10,2").Split(',');
                            var decimals = spec.Length > 1 ? int.Parse(spec[1]) : 0;
                            if (decimal.TryParse(number, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var numeric))
                            {
                                if (numeric.ToString("F" + decimals, CultureInfo.InvariantCulture).Length > column.Length)
                                    throw new ProcessingException("NUMBER_OVERFLOW", "dbf-write", "Число не помещается в поле DBF")
                                    { Field = column.Name, SourceRow = rowIndex, Actual = number };
                                if (decimal.Round(numeric, decimals) != numeric)
                                    Warn("NUMBER_PRECISION_LOSS", "Точность числа превышает точность поля DBF", rowIndex, column.Name, number);
                            }
                            record[i] = number;
                            break;
                        default:
                            throw new NotImplementedException($"Column {column.ColumnType} not supported yet!");
                    }
                    }
                    catch (ProcessingException) { throw; }
                    catch (Exception error)
                    {
                        throw new ProcessingException("DBF_FIELD_ERROR", "dbf-write", error.Message, error)
                        { Field = column.Name, SourceRow = rowIndex, Actual = rawValue };
                    }
                }

                db.Write(record, true);
                RecordsWritten++;
            }

            private void Warn(string code, string message, int row, string field, object actual)
            {
                owner.OnWarning?.Invoke(new DataWarning { Code = code, Message = message, SourceRow = row, Field = field, Actual = actual });
                if (owner.Strict) throw new ProcessingException(code, "dbf-write", message)
                { SourceRow = row, Field = field, Actual = actual };
            }

            public void Dispose()
            {
                db.Close();
            }
        }

    }
}
