using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services.Preload;
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
    public class PreloadService
    {
        private readonly ILogger logger;
        private readonly IUnityContainer container;
        private readonly ConfigProvider provider;
        private readonly IWebService web;

        public PConfig Settings => settings;
        private readonly PConfig settings;

        public PreloadService(ILogger logger, IUnityContainer container)
        {
            this.logger = logger;
            this.container = container;
            provider = container.Resolve<ConfigProvider>();
            web = container.Resolve<IWebService>();
            FileStorage.Load(Constants.PreloadFile, out settings, new PConfig());
            SaveChanges();
        }

        public void SaveChanges()
        {
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
                var url = await GetUrl(token);
                logger.Info("Загружаем внешний конфиг: " + url);
                var file = await web.GetFile(url, token);
                File.WriteAllText(Constants.SettingsFile, file);
                provider.ReloadConfig();
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

        private async Task<string> GetUrl(CancellationToken token)
        {
            if (settings.UseForceURL && !string.IsNullOrEmpty(settings.ForceURL)) return settings.ForceURL;
            var repo = await web.Get<Repository>(settings.Repository, token);

            var tag = repo.Tags.FirstOrDefault(x => x.Title == settings.Tag);
            if (tag == null) throw new InvalidOperationException($"Не найден тэг \"{settings.Tag}\" в репозитории!");

            if (!tag.Versions.TryGetValue(settings.Version, out var filename))
                throw new InvalidOperationException($"Не найдена версия \"{settings.Version}\" в тэге \"{tag.Title}\"!");

            return new URLBuilder().Append(repo.Root).Append(tag.Url).Append(filename).Build();

        }
    }
}
