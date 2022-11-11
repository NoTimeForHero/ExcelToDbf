using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Jint;
using Jint.Native;
using Newtonsoft.Json;

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

        public static Engine ClearValue(this Engine engine, JsValue name) =>
            // ReSharper disable once AssignNullToNotNullAttribute
            engine.SetValue(name, (object)null);

        // https://stackoverflow.com/a/51241629
        public static T[] GetRow<T>(this T[,] matrix, int rowNumber) =>
            Enumerable.Range(0, matrix.GetLength(1))
                .Select(x => x < 1 ? default : matrix[rowNumber, x])
                .ToArray();

        public static IEnumerable<T[]> AsRowArray<T>(this T[,] matrix)
        {
            var len = matrix.GetLength(0);
            for (int i = 1; i <= len; i++)
            {
                yield return matrix.GetRow(i);
            }
        }

        public static T Deserialize<T>(this Jint.Native.Json.JsonSerializer serializer, JsValue value)
        {
            var json = serializer.Serialize(value, Undefined.Instance, Undefined.Instance).AsString();
            return JsonConvert.DeserializeObject<T>(json);
        }
    }
}
