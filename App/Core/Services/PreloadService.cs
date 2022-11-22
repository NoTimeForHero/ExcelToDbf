using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelToDbf.Core.ViewModels;
using ExcelToDbf.Core.Views;
using NLog;
using Unity;

namespace ExcelToDbf.Core.Services
{
    internal class PreloadService
    {
        private readonly ILogger logger;
        private readonly IUnityContainer container;

        public PreloadService(ILogger logger, IUnityContainer container)
        {
            this.logger = logger;
            this.container = container;
        }

        public async Task RunGUI()
        {
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
                await Task.Delay(1000, token);
            }
            catch (TaskCanceledException)
            {
                logger.Warn("Загрузка конфига была отменена пользователем!");
            }
        }

        private class Config
        {
            public string Repository { get; set; }
            public string Tag { get; set; }
            public string Version { get; set; }
            public bool Enabled { get; set; }
        }
    }
}
