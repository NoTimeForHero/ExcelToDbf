using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Utils.Extensions;
using Jint;
using Jint.Native;
using Jint.Native.Array;
using Jint.Native.Function;
using Jint.Native.Json;
using Jint.Native.Object;

namespace ExcelToDbf.Core.Services.Scripts
{
    internal class DocForm
    {
        public string Name { get; set; }
        public ScriptFunctionInstance Rules { get; set; }
        public ScriptFunctionInstance Write { get; set; }
        public DbfFields[] Fields { get; set; }
        public XSettings Settings { get; set; }

        public class XSettings
        {
            public int StartY { get; set; }
            public int EndX { get; set; }
        }

        public class DbfFields
        {
            public string Name { get; set; }
            public string Type { get; set; }
            public string Length { get; set; }
        }
    }
}
