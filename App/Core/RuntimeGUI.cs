using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Core.ViewModels;
using ExcelToDbf.Core.Views;
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
        private readonly LastLaunchService lastLaunch;
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
            lastLaunch = new LastLaunchService(logger);
            model = container.Resolve<MainViewModel>();
            converter = container.Resolve<ConvertService>();
            currentState = UIState.SelectFiles;
        }

        public void Run()
        {
            UpdateUI(UIState.SelectFiles);

            var session = lastLaunch.Load();
            var fileSelector = container.Resolve<FileSelectorVM>();
            if (Directory.Exists(session.Path))
            {
                fileSelector.Path = session.Path;
            }

            if (session.Results?.Count > 0)
            {
                container.Resolve<ConvertResultVM>().Results = session.Results;
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
            var view = container.Resolve<MainView>();
            lastLaunch.TrackWindow(view, () => fileSelector.Path);
            view.ShowDialog();
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

            lastLaunch.SaveResults(container.Resolve<FileSelectorVM>().Path, results);

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
