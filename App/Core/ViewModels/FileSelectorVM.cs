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

        public string SelectedCount { [ObservableAsProperty] get; }

        public FileSelectorVM()
        {
        }

        public FileSelectorVM(FolderService service)
        {
            SelectPathCommand = ReactiveCommand.Create(() =>
            {
                Path = UniversalFolderSelector.ShowDialog() ?? Path;
            });

            CheckedCommand = ReactiveCommand.Create<string>((arg) =>
            {
                MessageBox.Show($"Checked all: " + (arg == "true" ? "YES" : "NO"));
            });

            this.WhenAnyValue(x => x.Path).Subscribe(service.Update);

            service.Connect()
                .ObserveOn(RxApp.MainThreadScheduler)
                .Bind(out _files)
                .DisposeMany()
                .Subscribe();

            service.Connect()
                .Select(x => x.Count.ToString())
                .ToPropertyEx(this, x => x.SelectedCount);
        }
    }
}
