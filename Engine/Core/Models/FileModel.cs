using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using DynamicData.Binding;
using ExcelToDbf.Utils;
using Newtonsoft.Json;
using ReactiveUI;

namespace ExcelToDbf.Core.Models
{
    [JsonObject(MemberSerialization.OptOut)]
    public class FileModel : ReactiveObject
    {
        public string FullPath { get; set; }
        public string FileName { get; set; }
        public long Size { get; set; }
        public DateTime Created { get; set; }

        private bool mustConvert;
        public bool MustConvert
        {
            get => mustConvert;
            set => this.RaiseAndSetIfChanged(ref mustConvert, value);
        }
    }
}
