using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Core.Models
{
    public readonly struct Point
    {
        public override string ToString()
        {
            return $"{{Point Y={Y}, X={X}}}";
        }

        public bool Equals(Point other)
        {
            return X == other.X && Y == other.Y;
        }

        public override bool Equals(object obj)
        {
            return obj is Point other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (X * 397) ^ Y;
            }
        }

        public int X { get; }
        public int Y { get; }

        public Point(int y, int x)
        {
            X = x;
            Y = y;
        }
    }
}
