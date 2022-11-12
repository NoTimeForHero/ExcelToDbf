using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Utils.Extensions;
using ExcelToDbf.Utils.Serializers;
using Jint;
using Jint.Native;
using Jint.Native.Array;
using Jint.Native.Function;
using Jint.Native.Json;
using Jint.Native.Object;
using Newtonsoft.Json;
using JsonSerializer = Newtonsoft.Json.JsonSerializer;

namespace ExcelToDbf.Core.Models
{
    [JsonConverter(typeof(DocFormConverter))]
    public class DocForm
    {
        public string Id { get; set; }
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
