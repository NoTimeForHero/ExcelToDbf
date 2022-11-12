using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
using ExcelToDbf.Core.Services.Scripts.Data;
using ExcelToDbf.Utils.Extensions;
using NLog;
using Unity;

namespace ExcelToDbf.Core.Services
{
    internal class ConvertService
    {
        private readonly ILogger logger;
        private readonly ScriptEngine engine;
        private readonly DBFService database;
        private readonly Random random = new Random(100);
        private readonly Lazy<ExcelService> lazyExcel;

        public ConvertService(ILogger logger, IUnityContainer container)
        {
            this.logger = logger;
            engine = container.Resolve<ScriptEngine>();
            database = container.Resolve<DBFService>();
            lazyExcel = new Lazy<ExcelService>(() => container.Resolve<ExcelService>());
        }

        public ConvertProgress Progress { get; } = new ConvertProgress();


        public Task<List<Result>> Run(IEnumerable<FileModel> filesToConvert)
        {
            var files = filesToConvert.ToList();
            logger.Info($"Запущен процесс конвертации {files.Count} файлов!");
            logger.Debug($"Файлы: " +
                         files.Select(x => $"\"{x.FileName}\"").JoinString(", ")
                         );
            var task = Task.Factory.StartNew(() => RunInternal(files), TaskCreationOptions.LongRunning).Unwrap();
            return task;

        }

        private async Task<List<Result>> RunInternal(List<FileModel> input)
        {
            var files = input.ToList();
            var filesTotal = files.Count;
            var results = new List<Result>();

            Progress.Reset();
            Progress.GlobalInitialize(filesTotal, "Ожидание загрузки Excel...");

            var folderCtx = engine.Resolve<ConfigContext>();

            foreach (var (file, curFile) in files.WithIndex())
            {
                var result = new Result { File = file };
                results.Add(result);
                var filename = file.FileName;
                logger.Info($"Конвертация файла: {filename}");
                Progress.FileInitialize(curFile+1, filename);
                try
                {
                    var outputFile = folderCtx.GetOutputFilename(file);
                    result.OutputFilename = outputFile;
                    await ProcessFile(ref result, file, outputFile);
                }
                catch (Exception ex)
                {
                    logger.Warn($"Ошибка обработки файла: {filename}");
                    logger.Warn(ex);
                    result.Error = $"{ex.GetType().Name}: {ex.Message}";
                }
            }
            return results;
        }

        private Task ProcessFile(ref Result result, FileModel file, string outputFile)
        {
            Progress.LocalText = $"Поиск подходящих форм для файла: {file.FileName}";

            var excel = lazyExcel.Value;
            excel.OpenWorksheet(file.FullPath);

            var context = engine.Resolve<ExcelContext>().Connect(excel.GetCellValue);
            var search = context.SearchForm(file);
            var form = search.Result;

            result.SearchResult = search;
            if (form == null)
            {
                logger.Warn($"Для файла \"{file.FileName}\" не найдено подходящих форм обработки!");
                return Task.CompletedTask;
            }

            using (var dbfFile = database.Make(form, outputFile))
            {
                dbfFile.WriteHeaders();

                var totalRows = excel.SheetRows;
                Progress.DocumentTotal = totalRows;
                int currentRow = form.Settings.StartY;
                try
                {
                    foreach (var range in excel.IterateRanges(form.Settings.StartY, form.Settings.EndX))
                    {
                        foreach (var record in range.AsRowArray())
                        {
                            currentRow++;
                            Progress.SetProgress(currentRow, totalRows, $"Обработка массива строк");
                            var transformed = context.Transform(form, record);
                            dbfFile.WriteRecord(transformed);
                        }
                    }
                }
                catch (StopFunctionException)
                {
                    logger.Info($"Обработка файла была завершена JS условием на {currentRow} строке!");
                }
                result.Converted = true;
            }
            return Task.CompletedTask;
        }

        public class Result
        {
            public bool Converted { get; set; }
            public SearchFormResult SearchResult { get; set; }
            public FileModel File { get; set; }
            public string OutputFilename { get; set; }
            public string Error { get; set; }
        }
    }
}
