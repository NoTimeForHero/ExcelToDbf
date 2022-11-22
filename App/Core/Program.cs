using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
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
            container.RegisterSingleton<PreloadService>();
            container.RegisterSingletonMVVM<MainView, MainViewModel>();
            container.RegisterSingletonMVVM<FileSelectorView, FileSelectorVM>();
            container.RegisterSingletonMVVM<ConvertResultView, ConvertResultVM>();
            container.RegisterSingletonMVVM<ProgressView, ProgressVM>();
            container.RegisterSingletonMVVM<LoadingView, LoadingVM>();
            container.Resolve<ScriptEngine>()
                .Register<GenericContext>()
                .Register<ConfigContext>()
                .Register<ExcelContext>()
                .Register<IConfigContext, ConfigContext>();
            container.RegisterFactory<ConfigProvider>(x => x.Resolve<ScriptEngine>().Resolve<ConfigContext>().Data);
            container.RegisterInstance(this);
        }

        private void Debug()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\ExcelToDbf\Data");
            container.Resolve<FileSelectorVM>().Path = path;
            container.Resolve<FolderService>().SelectAll(false);
            container.Resolve<FolderService>().SelectWhere(x => x.FileName == "Example2.xlsx", true);
        }

        public async Task Run(string[] args)
        {
            try
            {
                LogManager.ThrowExceptions = true;
                LogManager.ThrowConfigExceptions = true;
                var logger = LogManager.GetCurrentClassLogger();
                logger.Info("Приложение было запущено");

                Debug();
                await preload.RunGUI();
                var gui = new RuntimeGUI(container, logger);
                gui.Run();
                container.Dispose();
                Application.Current.Shutdown();
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
