using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.ViewModels
{
    public class MainViewModel : ReactiveObject
    {
        [Reactive]
        public string Title { get; set; } = "Конвертирование Excel документов в DBF";

        [Reactive]
        public string HeaderTitle { get; set; } = "Название фирмы";

        [Reactive]
        public string HeaderDescription { get; set; } = "Дополнительная информация";

        [Reactive]
        public string ButtonActionTitle { get; set; } = "Действие";

        [Reactive]
        public bool ButtonActionEnabled { get; set; } = true;

        public ReactiveCommand<Unit, Unit> ActionCommand { get; set; }

        [Reactive]
        public FrameworkElement ViewBody { get; set; } = null;

        public MainViewModel()
        {
            ActionCommand = ReactiveCommand.CreateFromTask(() =>
            {
                MessageBox.Show("Кнопка нажата?");
                return Task.CompletedTask;
            }, canExecute: this.WhenAnyValue(x => x.ButtonActionEnabled));
        }
    }
}
