using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using DynamicData;
using DynamicData.Aggregation;
using DynamicData.Binding;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services;
using ExcelToDbf.Utils;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.Core.ViewModels
{
    internal class FileSelectorVM : ReactiveObject
    {
        [Reactive]
        public string Path { get; set; } = "";

        public ReadOnlyObservableCollection<FileModel> _files;
        public ReadOnlyObservableCollection<FileModel> Files => _files;

        public ReactiveCommand<Unit, Unit> SelectPathCommand { get; }
        public ReactiveCommand<string, Unit> CheckedCommand { get; }

        public int SelectedCount { [ObservableAsProperty] get; }

        //[Reactive]
        //public int SelectedCount { get; set; } = -1;


        public FileSelectorVM()
        {
        }

        private void Generate(SourceList<FileModel> list, int count, bool? isChecked = null)
        {
            var random = new Random();
            list.Clear();
            list.Edit(innerList =>
            {
                foreach (var index in Enumerable.Range(1, count))
                {
                    var chk = isChecked ?? random.Next(100) > 50;
                    list.Add(new FileModel { FileName = $"File #{index}", MustConvert = chk });
                }
            });
        }

        // TODO: Добавить IDisposable
        public FileSelectorVM(FolderService service)
        {
            SelectPathCommand = ReactiveCommand.Create(() =>
            {
                Path = UniversalFolderSelector.ShowDialog() ?? Path;
            });

            CheckedCommand = ReactiveCommand.Create<string>((arg) =>
            {
                service.SelectAll(arg == "true");
            });

            this.WhenAnyValue(x => x.Path).Subscribe(service.Update);

            service.Connect()
                //.Sort(SortExpressionComparer<FileModel>.Ascending(vm => vm.Size))
                .ObserveOn(RxApp.MainThreadScheduler)
                //.ObserveOnDispatcher()
                .Bind(out _files)
                //.DisposeMany()
                .Subscribe();

            service
                .Connect()
                .AutoRefresh(x => x.MustConvert)
                .Filter(x => x.MustConvert)
                .Count()
                .ToPropertyEx(this, x => x.SelectedCount);
        }
    }
}
