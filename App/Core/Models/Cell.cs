using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Core.Models
{
    public struct Cell
    {
        public int Y { get; set; }
        public int X { get; set; }
        public string Value { get; set; }
    }
}
