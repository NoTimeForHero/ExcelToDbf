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
using ExcelToDbf.Core.ViewModels;
using ExcelToDbf.Core.Views;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using ReactiveUI;
using Unity;

namespace ExcelToDbf.Core
{
    internal class RuntimeGUI
    {
        private readonly IUnityContainer container;
        private readonly MainViewModel model;
        private readonly ConvertService converter;
        private UIState currentState;

        private enum UIState
        {
            SelectFiles,
            DisplayLogs,
            Converting
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

        public RuntimeGUI(IUnityContainer container)
        {
            this.container = container;
            model = container.Resolve<MainViewModel>();
            converter = container.Resolve<ConvertService>();
            currentState = UIState.SelectFiles;
        }

        public void Run()
        {
            UpdateUI(UIState.SelectFiles);

            if (File.Exists(Constants.LastLaunchFile))
            {
                var results = JsonConvert.DeserializeObject<List<ConvertService.Result>>(File.ReadAllText(Constants.LastLaunchFile));
                var vm = container.Resolve<ConvertResultVM>();
                Console.WriteLine("RESULTS!");
                UpdateUI(UIState.DisplayLogs);
            }

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

            var result = await converter.Run(files);
            File.WriteAllText(Constants.LastLaunchFile, JsonConvert.SerializeObject(result, Formatting.Indented));

            UpdateUI(UIState.DisplayLogs);
        }
    }
}
