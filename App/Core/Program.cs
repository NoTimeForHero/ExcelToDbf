using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ExcelToDbf.Extensions;
using ExcelToDbf.ViewModels;
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
            container.RegisterSingletonMVVM<MainView, MainViewModel>();
            container.RegisterInstance(this);
        }

        private void PrepareActions(MainViewModel model)
        {
            var body = new Label();
            body.Content = "Ожидание пользовательского ввода...";
            body.FontSize = 36;
            body.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            body.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            model.ViewBody = body;

            model.ActionCommand = ReactiveCommand.CreateFromTask(async () =>
            {
                model.ButtonActionEnabled = false;
                body.Content = "Загрузка данных...";
                await Task.Delay(2000);
                model.ButtonActionEnabled = true;
                model.ButtonActionTitle = "Повторить";
                body.Content = "Данные успешно загружены!";
            });
        }

        public void Run(string[] args)
        {
            var model = container.Resolve<MainViewModel>();
            model.HeaderTitle = "ООО \"Рога и копыта\"";
            model.HeaderDescription = "Версия 0.0.0.1 Альфа\nЗагружено 24 форм\nПоследнее обновление 01.01.2020";
            model.Title = "Конвертирование Excel документов в DBF (версия 3.0.0.1 Альфа)";
            model.ButtonActionTitle = "Конвертировать!";
            PrepareActions(model);

            var form = container.Resolve<MainView>();
            form.Show();
            //
            // var model = new MainViewModel();
            // model.Name = "John";
            // model.LastName = "Doe";
            //
            // var form = new MainView(model);
            // form.Show();
        }

    }
}
