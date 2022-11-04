using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils;
using Jint;
using NLog;

namespace ExcelToDbf.Core.Services.Scripts.Context
{
    internal class FolderContext : AbstractContext
    {
        private readonly ILogger logger;

        public FolderContext(ILogger logger, Engine engine) : base(engine)
        {
            this.logger = logger;
        }

        public string GetOutputFilename(FileModel file)
        {
            DirectoryInfo dir = new DirectoryInfo(file.FullPath);
            PathHelper helper = new PathHelper(dir);
            engine.SetValue("dir", (Func<int, string>)helper.GetLevel);
            engine.SetValue("dirCount", helper.Count);
            var outputName = engine
                .SetValue("file", Path.GetFileNameWithoutExtension(file.FileName))
                .Evaluate("app.getOutputFilename(file)")
                .AsString();
            var baseDir = Path.GetDirectoryName(file.FullPath)
                          ?? throw new Exception("Directory not found!");
            logger.Debug($"Преобразование имени \"{file.FileName}\" в \"{outputName}\"");
            return Path.Combine(baseDir, outputName);
        }
    }
}
