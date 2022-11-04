﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Utils.Extensions;
using NLog;
using Unity;

namespace ExcelToDbf.Core.Services
{
    internal class ConvertService
    {
        private readonly ILogger logger;
        private readonly ScriptEngine engine;
        private readonly Random random = new Random(100);
        private readonly Lazy<ExcelService> lazyExcel;

        public ConvertService(ILogger logger, IUnityContainer container)
        {
            this.logger = logger;
            engine = container.Resolve<ScriptEngine>();
            lazyExcel = new Lazy<ExcelService>(() => container.Resolve<ExcelService>());
        }

        public ConvertProgress Progress { get; } = new ConvertProgress();


        public async Task Run(IEnumerable<FileModel> filesToConvert)
        {
            var files = filesToConvert.ToList();
            logger.Debug($"Набор {files.Count} файлов для конвертации: " +
                         files.Select(x => $"\"{x.FileName}\"").JoinString(", ")
                         );
            var task = Task.Factory.StartNew(() => RunInternal(files), TaskCreationOptions.LongRunning).Unwrap();
            await task;

        }

        private async Task RunInternal(List<FileModel> input)
        {
            var files = input.ToList();
            var filesTotal = files.Count;

            Progress.Reset();
            Progress.GlobalInitialize(filesTotal, "Ожидание загрузки Excel...");

            foreach (var (file, curFile) in files.WithIndex())
            {
                var filename = file.FileName;
                logger.Info($"Конвертация файла: {filename}");
                Progress.FileInitialize(curFile+1, filename);
                try
                {
                    var outputFile = engine.GetOutputFilename(file);
                    await ProcessFile(file.FullPath, outputFile);
                }
                catch (Exception ex)
                {
                    logger.Warn($"Ошибка обработки файла: {filename}");
                    logger.Warn(ex);
                }
            }
        }

        private Task ProcessFile(string inputFile, string outputFile)
        {
            var excel = lazyExcel.Value;
            excel.OpenWorksheet(inputFile);

            // engine.Excel.FindForm(excel.worksheet);

            // Progress.DocumentTotal = rows;
            // int curRow = 0;
            // while (curRow < rows)
            // {
            //     Progress.SetProgress(curRow, rows, $"Обработка массива строк");
            //     curRow += random.Next(100, 400);
            //     Thread.Sleep(20);
            //     //await random.Delay(5, 50);
            // }
            // File.WriteAllText(outputFile, "Test");
            return Task.CompletedTask;
        }
    }
}
