using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Utils;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.Core.ViewModels
{
    internal class FileSelectorVM : ReactiveObject
    {
        [Reactive]
        public string Path { get; set; } = "";

        public ReactiveCommand<Unit, Unit> SelectPathCommand { get; }

        public FileSelectorVM()
        {
        }

        public FileSelectorVM(Config config)
        {
            SelectPathCommand = ReactiveCommand.Create(() =>
            {
                Path = UniversalFolderSelector.ShowDialog() ?? Path;
            });
        }
    }
}
