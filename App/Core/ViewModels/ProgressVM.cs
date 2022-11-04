using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services;
using ReactiveUI;

namespace ExcelToDbf.Core.ViewModels
{
    internal class ProgressVM : ReactiveObject
    {

        public ConvertProgress Progress { get; set; }

        public ProgressVM()
        {
            Progress = new ConvertProgress
            {
                DocumentCurrent = 500,
                DocumentTotal = 10000,
                FilesCurrent = 5,
                FilesTotal = 10,
                GlobalText = "Глобальная строка состояния",
                LocalText = "Локальная строка состояния",
            };
        }

        public ProgressVM(ConvertService service)
        {
            Progress = service.Progress;
        }

    }
}
