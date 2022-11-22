using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services.Scripts;
using ExcelToDbf.Core.Services.Scripts.Context;
using ExcelToDbf.Core.ViewModels;
using ExcelToDbf.Core.Views;
using ExcelToDbf.Utils;
using Newtonsoft.Json;
using NLog;
using Unity;
using PConfig = ExcelToDbf.Core.Services.Preload.Config;

namespace ExcelToDbf.Core.Services
{
    internal class PreloadService
    {
        private readonly ILogger logger;
        private readonly IUnityContainer container;
        private readonly IWebService web;

        private readonly PConfig settings;

        public PreloadService(ILogger logger, IUnityContainer container)
        {
            this.logger = logger;
            this.container = container;
            web = container.Resolve<IWebService>();
            FileStorage.Load(Constants.PreloadFile, out settings);
            settings = settings ?? new PConfig();
            FileStorage.Save(Constants.PreloadFile, settings);
        }

        public async Task RunGUI()
        {
            if (!settings.Enabled) return;
            var status = container.Resolve<LoadingVM>();
            status.MainText = "Загрузка форм...";
            var tsCancel = new CancellationTokenSource();
            var view = container.Resolve<LoadingView>();
            view.Closed += (o, ev) => tsCancel.Cancel();
            view.Show();
            await Load(tsCancel.Token);
            view.Close();
        }

        private async Task Load(CancellationToken token)
        {
            try
            {
                var url = settings.ForceURL;
                var file = await web.GetFile(url);
                File.WriteAllText(Constants.SettingsFile, file);
                container.Resolve<ScriptEngine>().Resolve<ConfigContext>().ReloadConfig();
                // await Task.Delay(1000, token);
            }
            catch (TaskCanceledException)
            {
                logger.Warn("Загрузка конфига была отменена пользователем!");
            }
            catch (Exception ex)
            {
                logger.Warn("Не удалось загрузить конфигурацию!");
                logger.Warn(ex);
            }
        }
    }
}
