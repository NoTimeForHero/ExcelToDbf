﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DynamicData;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Utils;
using NLog;

namespace ExcelToDbf.Core.Services
{
    internal class FolderService
    {
        private readonly ILogger logger;
        private readonly ConfigProvider pvConfig;
        private readonly SourceList<FileModel> _files;

        public IObservable<IChangeSet<FileModel>> Connect() => _files.Connect();

        public FolderService(ILogger logger, ConfigProvider pvConfig)
        {
            _files = new SourceList<FileModel>();
            this.logger = logger;
            this.pvConfig = pvConfig;
        }

        public void SelectWhere(Func<FileModel,bool> predicate, bool isChecked)
        {
            _files.Edit(list =>
            {
                foreach (var item in list)
                {
                    if (predicate(item)) item.MustConvert = isChecked;
                }
            });
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
            logger.Info($"Пользователем выбрана директория: {path}");
            var range = DirectoryUtils.GetFilesByExtension(path, pvConfig.Config.Extensions)
                .Select(x => new FileModel
                {
                    MustConvert = true,
                    FileName = x.FileName,
                    FullPath = x.FullPath,
                    Created = x.Created,
                    Size = x.Size,
                }).ToList();
            logger.Info($"Файлов было найдено: {range.Count}");
            _files.Clear();
            _files.AddRange(range);
        }
    }
}
