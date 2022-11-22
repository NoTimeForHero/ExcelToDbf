using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.Core.ViewModels
{
    public class LoadingVM : ReactiveObject
    {
        [Reactive]
        public string MainText { get; set; }

        [Reactive]
        public string SecondaryText { get; set; }
    }
}
