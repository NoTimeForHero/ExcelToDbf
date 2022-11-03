using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Core.ViewModels;
using ExcelToDbf.Core.Views;
using ExcelToDbf.Utils.Extensions;
using ReactiveUI;
using Unity;
using Unity.NLog;

namespace ExcelToDbf.Core
{
    internal class Program
    {
        private readonly IUnityContainer container;

        public Program()
        {
            container = new UnityContainer();
            container.AddNewExtension<NLogExtension>();
            container.RegisterSingleton<ScriptEngine>();
            container.RegisterSingleton<FolderService>();
            container.RegisterSingletonMVVM<MainView, MainViewModel>();
            container.RegisterSingletonMVVM<FileSelectorView, FileSelectorVM>();
            container.RegisterFactory<Config>((u) => u.Resolve<ScriptEngine>().GetConfig());
            container.RegisterInstance(this);
        }


        public void Run(string[] args)
        {
            var model = container.Resolve<MainViewModel>();
            model.ChildVM = container.Resolve<FileSelectorVM>();

            model.ActionButton.Title = "Конвертировать";
            model.ActionButton.Image = MainViewModel.RActionButton.ImageType.Settings;

            model.ActionButton.Command = ReactiveCommand.CreateFromTask(async() =>
            {
                model.ActionButton.Visible = false;
                await Task.Delay(3000);
                model.ActionButton.Visible = true;
            });

            container.Resolve<MainView>().Show();
        }

    }
}
