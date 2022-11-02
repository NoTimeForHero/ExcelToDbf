using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Utils;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.Core.Models
{
    internal class FileModel : ReactiveObject
    {
        public string FullPath { get; set; }
        public string FileName { get; set; }
        public string Size { get; set; }
        public DateTime Created { get; set; }

        [Reactive]
        public bool Selected { get; set; }
    }
}
