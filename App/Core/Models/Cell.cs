using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Core.Models
{
    public struct Cell
    {
        public Cell(int y, int x, string value = null)
        {
            Y = y;
            X = x;
            Value = value;
        }

        public int Y { get; set; }
        public int X { get; set; }
        public string Value { get; set; }
    }
}
