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

        public void SelectAll(bool isChecked)
        {
            _files.Edit(list =>
            {
                foreach (var item in list)
                {
                    item.MustConvert = isChecked;
                }
            });
        }

        public IEnumerable<FileModel> GetFiles(bool? isChecked = false)
        {
            if (isChecked == null) return _files.Items;
            // ReSharper disable once PossibleInvalidOperationException
            return _files.Items.Where(file => file.MustConvert == isChecked.Value);
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
