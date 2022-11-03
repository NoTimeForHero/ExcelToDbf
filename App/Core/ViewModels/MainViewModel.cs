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

namespace ExcelToDbf.Core.ViewModels
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
        public RActionButton ActionButton { get; set; } = new RActionButton();

        [Reactive]
        public ReactiveObject ChildVM { get; set; } = null;

        public MainViewModel()
        {
        }

        public MainViewModel(Config config)
        {
            HeaderTitle = config.Header.Title;
            HeaderDescription = config.Header.Status;
        }


        public class RActionButton : ReactiveObject
        {
            [Reactive]
            public string Title { get; set; } = "Действие";

            [Reactive]
            public bool Enabled { get; set; } = true;

            [Reactive]
            public bool Visible { get; set; } = true;

            [Reactive]
            public ImageType Image { get; set; } = ImageType.None;

            public ReactiveCommand<Unit, Unit> Command { get; set; }

            public RActionButton()
            {
                Command = ReactiveCommand.CreateFromTask(() =>
                {
                    MessageBox.Show("Кнопка нажата?");
                    return Task.CompletedTask;
                }, canExecute: this.WhenAnyValue(x => x.Enabled));
            }

            public enum ImageType
            {
                None,
                Settings,
                Folder
            }
        }
    }
}
