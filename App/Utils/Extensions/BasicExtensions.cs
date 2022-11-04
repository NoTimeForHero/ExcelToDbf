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

        public static string JoinString(this IEnumerable<string> values, string separator)
            => string.Join(separator, values);

        public static string NestedMessages(this Exception ex)
        {
            var currentEx = ex;
            var builder = new StringBuilder();
            var level = 0;
            while (true)
            {
                var message = level == 0 ? "Ошибка" : $"Вложенная ошибка {level}";
                builder.Append($"\n[{message}]: ");
                builder.Append(currentEx.Message);
                if (currentEx.InnerException == null) break;
                currentEx = currentEx.InnerException;
                level++;
            }
            return builder.ToString();
        }
    }
}
