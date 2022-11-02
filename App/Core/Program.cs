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
            model.ButtonActionTitle = "Конвертировать!";
            model.ChildVM = container.Resolve<FileSelectorVM>();

            container.Resolve<MainView>().Show();
        }

    }
}
