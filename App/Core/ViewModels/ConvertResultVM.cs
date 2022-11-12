using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.Core.ViewModels
{
    internal class ConvertResultVM : ReactiveObject
    {
        [Reactive]
        public List<ConvertService.Result> Results { get; set; }

        public ConvertResultVM()
        {

        }
    }
}
