using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DynamicData;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils;

namespace ExcelToDbf.Core.Services
{
    internal class FolderService
    {
        private readonly Config config;
        private readonly SourceList<FileModel> _files;

        public IObservable<IChangeSet<FileModel>> Connect() => _files.Connect();

        public FolderService(Config config)
        {
            _files = new SourceList<FileModel>();
            this.config = config;
        }



        public void Update(string path)
        {
            var range = DirectoryUtils.GetFilesByExtension(path, config.Extensions)
                .Select(x => new FileModel
                {
                    FileName = x.FileName,
                    FullPath = x.FullPath,
                    Created = x.Created,
                    Size = x.Size,
                });
            _files.Clear();
            _files.AddRange(range);
        }
    }
}
