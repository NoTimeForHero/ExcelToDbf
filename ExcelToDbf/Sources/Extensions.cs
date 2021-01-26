using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Jint.Native;
using Jint.Native.Object;

namespace ExcelToDbf.Sources
{
    public static class JsExt
    {
        public static JsValue getOrDefault(this ObjectInstance jObj, string propertyName, JsValue orDefault)
        {
            var value = jObj.Get(propertyName);
            if (value.IsNull() || value.IsUndefined()) return orDefault;
            return value;
        }
    }

    public static class ArrayExt
    {
        // https://stackoverflow.com/questions/27427527/how-to-get-a-complete-row-or-column-from-2d-array-in-c-sharp
        public static T[] GetRow<T>(this T[,] matrix, int rowNumber, int start=0)
        {
            return Enumerable.Range(start, matrix.GetLength(1))
                .Select(x => matrix[rowNumber, x])
                .ToArray();
        }

    }
}
