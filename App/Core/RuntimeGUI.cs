using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
using ExcelToDbf.Core.ViewModels;
using ExcelToDbf.Core.Views;
using ExcelToDbf.Utils;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NLog;
using ReactiveUI;
using Unity;

namespace ExcelToDbf.Core
{
    internal class RuntimeGUI
    {
        private readonly IUnityContainer container;
        private readonly MainViewModel model;
        private readonly ConvertService converter;
        private readonly ILogger logger;
        private UIState currentState;

        private enum UIState
        {
            SelectFiles,
            DisplayLogs,
            Converting
        }

        public RuntimeGUI(IUnityContainer container, ILogger logger)
        {
            this.container = container;
            this.logger = logger;
            model = container.Resolve<MainViewModel>();
            converter = container.Resolve<ConvertService>();
            currentState = UIState.SelectFiles;
        }

        public void Run()
        {
            UpdateUI(UIState.SelectFiles);

            if (FileStorage.Load<LastLaunch>(Constants.LastLaunchFile, out var lastLaunch))
            {
                container.Resolve<ConvertResultVM>().Results = lastLaunch.Results;
                container.Resolve<FileSelectorVM>().Path = lastLaunch.Path;
                UpdateUI(UIState.DisplayLogs);
            }

            model.CommandSettings = ReactiveCommand.Create(() =>
            {
                container.Resolve<EditPreloadView>().ShowDialog();
            });

            model.ActionButton.Command = ReactiveCommand.CreateFromTask(async () =>
            {
                switch (currentState)
                {
                    case UIState.SelectFiles:
                        await Convert();
                        break;
                    case UIState.DisplayLogs:
                        UpdateUI(UIState.SelectFiles);
                        break;
                    default:
                        throw new NotImplementedException(currentState.ToString());
                }
            });
            container.Resolve<MainView>().ShowDialog();
        }

        private async Task Convert()
        {
            var files = container.Resolve<FolderService>().GetFiles(true).ToList();

            if (files.Count == 0)
            {
                var message = "Вы не выбрали ни одного файла для конвертации!\nДля продолжения выберите хотя бы один файл!";
                MessageBox.Show(message, "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            UpdateUI(UIState.Converting);

            var results = await converter.Run(files);
            container.Resolve<ConvertResultVM>().Results = results;

            var lastLaunch = new LastLaunch
            {
                Path = container.Resolve<FileSelectorVM>().Path,
                Results = results
            };
            FileStorage.Save(Constants.LastLaunchFile, lastLaunch);

            UpdateUI(UIState.DisplayLogs);
        }


        private void UpdateUI(UIState? newState)
        {
            if (newState.HasValue) currentState = newState.Value;
            switch (currentState)
            {
                case UIState.Converting:
                    model.ActionButton.Visible = false;
                    model.CloseConfirmation = true;
                    model.ChildVM = container.Resolve<ProgressVM>();
                    break;
                case UIState.SelectFiles:
                    model.ActionButton.Visible = true;
                    model.CloseConfirmation = false;
                    model.ChildVM = container.Resolve<FileSelectorVM>();
                    model.ActionButton.Title = "Конвертировать";
                    model.ActionButton.Image = MainViewModel.RActionButton.ImageType.Settings;
                    break;
                case UIState.DisplayLogs:
                    model.ActionButton.Visible = true;
                    model.CloseConfirmation = false;
                    model.ChildVM = container.Resolve<ConvertResultVM>();
                    model.ActionButton.Title = "Выбор файлов";
                    model.ActionButton.Image = MainViewModel.RActionButton.ImageType.Folder;
                    break;
                default:
                    throw new NotImplementedException(currentState.ToString());
            }
        }
    }
}
