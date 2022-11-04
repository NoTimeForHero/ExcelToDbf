using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils.Extensions;
using NLog;

namespace ExcelToDbf.Core.Services
{
    internal class ConvertService
    {
        private ILogger logger;
        private Config config;

        public ConvertService(ILogger logger, Config config)
        {
            this.logger = logger;
            this.config = config;
        }

        public ConvertProgress Progress { get; } = new ConvertProgress();


        public async Task Run(IEnumerable<FileModel> filesToConvert)
        {
            var files = filesToConvert.ToList();
            logger.Debug($"Набор {files.Count} файлов для конвертации: " +
                         files.Select(x => $"\"{x.FileName}\"").JoinString(", ")
                         );
            //await Task.Factory.StartNew(() => DemoRun(filesToConvert), TaskCreationOptions.LongRunning);
            await DemoRun(files);
        }

        private async Task DemoRun(List<FileModel> input)
        {
            var random = new Random(100);

            var files = input.ToList();
            var filesTotal = files.Count;

            Progress.Reset();
            Progress.GlobalInitialize(filesTotal);

            //var files = new 

            foreach (var (file, curFile) in files.WithIndex())
            {
                var filename = file.FileName;
                logger.Info($"Конвертация файла: {filename}");

                Progress.FileInitialize(
                    curFile+1,
                    $"Обработка файла: {filename}",
                    $"Открытие файла: {filename}"
                );

                await random.Delay(500, 3000);

                var rows = random.Next(5000, 60000);
                Progress.DocumentTotal = rows;
                int curRow = 0;
                while (curRow < rows)
                {
                    Progress.SetProgress(curRow, rows, $"Обработка массива строк");
                    curRow += random.Next(100, 400);
                    await random.Delay(5, 50);
                }
            }
        }
    }
}
