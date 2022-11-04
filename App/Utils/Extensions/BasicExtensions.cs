using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Utils.Extensions
{
    internal static class BasicExtensions
    {
        public static Task Delay(this Random rnd, int min = 1000, int max = 5000)
            => Task.Delay(rnd.Next(min, max));

        public static IEnumerable<(T item, int index)> WithIndex<T>(this IEnumerable<T> self)
            => self.Select((item, index) => (item, index));
    }
}
