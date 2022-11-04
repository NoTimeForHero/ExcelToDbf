using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Core.ViewModels;
using ExcelToDbf.Core.Views;
using ReactiveUI;
using Unity;

namespace ExcelToDbf.Core
{
    internal class RuntimeGUI
    {
        private readonly IUnityContainer container;

        public RuntimeGUI(IUnityContainer container)
        {
            this.container = container;
        }

        public void Run()
        {
            var model = container.Resolve<MainViewModel>();
            var converter = container.Resolve<ConvertService>();

            model.ChildVM = container.Resolve<FileSelectorVM>();

            model.ActionButton.Title = "Конвертировать";
            model.ActionButton.Image = MainViewModel.RActionButton.ImageType.Settings;

            model.ActionButton.Command = ReactiveCommand.CreateFromTask(async () =>
            {
                var files = container.Resolve<FolderService>().GetFiles(true).ToList();

                if (files.Count == 0)
                {
                    var message = "Вы не выбрали ни одного файла для конвертации!\nДля продолжения выберите хотя бы один файл!";
                    MessageBox.Show(message, "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                model.ActionButton.Visible = false;
                model.CloseConfirmation = true;
                model.ChildVM = container.Resolve<ProgressVM>();

                await converter.Run(files);

                model.ChildVM = container.Resolve<FileSelectorVM>();
                model.ActionButton.Visible = true;
                model.CloseConfirmation = false;
            });
            container.Resolve<MainView>().ShowDialog();
        }

    }
}
