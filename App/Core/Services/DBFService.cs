using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using Microsoft.Office.Interop.Excel;
using NLog;
using SocialExplorer.IO.FastDBF;

namespace ExcelToDbf.Core.Services
{
    internal class DBFService
    {
        protected readonly Config config;
        protected readonly ILogger logger;

        public DBFService(ILogger logger, Config config)
        {
            this.logger = logger;
            this.config = config;
        }

        public Work Make(DocForm form, string outputFilename) => new Work(this, form, outputFilename);

        public class Work : IDisposable
        {
            private readonly DbfFile oDBF;
            private readonly ILogger logger;
            private readonly DocForm form;

            internal Work(DBFService owner, DocForm form, string outputFilename)
            {
                this.form = form;
                logger = owner.logger;
                var config = owner.config;

                var encoding = Encoding.GetEncoding(config.System.OutputEncoding);

                oDBF = new DbfFile(encoding);
                oDBF.Open(outputFilename, FileMode.Create); // FileMode.Create = файл будет перезаписан если уже существует
                logger.Info($"Создаём DBF по пути: {outputFilename}");
                logger.Debug("и кодировкой: {encoding}");
            }

            public Work WriteHeaders()
            {
                foreach (var field in form.Fields)
                {
                    DbfColumn.DbfColumnType column;
                    string length = field.Length;
                    var type = field.Type?.ToLower() ?? "string";
                    switch (type)
                    {
                        case "date":
                            column = DbfColumn.DbfColumnType.Date;
                            length = length ?? "8";
                            break;
                        case "number":
                            column = DbfColumn.DbfColumnType.Number;
                            length = length ?? "10,2";
                            break;
                        case "string":
                            column = DbfColumn.DbfColumnType.Character;
                            break;
                        default:
                            throw new InvalidOperationException($"Unknown DBF field type: {field.Type}");
                    }
                    var lenParts = length.Split(',');
                    int nLen = int.Parse(lenParts[0]);
                    int nDec = lenParts.Length > 1 ? int.Parse(lenParts[1]) : 0;
                    oDBF.Header.AddColumn(new DbfColumn(field.Name, column, nLen, nDec));
                    logger.Debug($"Записываем поле '{type}' типа '{type}' длиной {nLen},{nDec}");
                }

                oDBF.WriteHeader();

                // config.System.OutputEncoding
                return this;
            }

            public void Dispose()
            {
                oDBF.Close();
            }
        }

    }
}
