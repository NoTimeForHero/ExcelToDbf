using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services;

namespace ExcelToDbf.Core.Models
{
    internal class LastLaunch
    {
        public string Path { get; set; }

        public List<ConvertService.Result> Results { get; set; }
    }
}
