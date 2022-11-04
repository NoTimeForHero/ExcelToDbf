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
using NLog;
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
            container.RegisterSingleton<ConvertService>();
            container.RegisterSingleton<ExcelService>();
            container.RegisterSingletonMVVM<MainView, MainViewModel>();
            container.RegisterSingletonMVVM<FileSelectorView, FileSelectorVM>();
            container.RegisterSingletonMVVM<ProgressView, ProgressVM>();
            container.RegisterFactory<Config>((u) => u.Resolve<ScriptEngine>().GetConfig());
            container.RegisterInstance(this);
        }

        public void Run(string[] args)
        {
            try
            {
                LogManager.ThrowExceptions = true;
                LogManager.ThrowConfigExceptions = true;
                var logger = LogManager.GetCurrentClassLogger();
                logger.Info("Приложение было запущено");

                new RuntimeGUI(container).Run();
                container.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.NestedMessages(),
                    "Критическая ошибка!",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                LogManager.GetCurrentClassLogger().Error(ex);
            }
        }

    }
}
