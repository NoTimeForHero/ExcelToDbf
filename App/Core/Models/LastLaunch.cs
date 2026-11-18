using System.Collections.Generic;
using ExcelToDbf.Core.Services;
using Newtonsoft.Json;

namespace ExcelToDbf.Core.Models
{
    internal class LastLaunch
    {
        public string Path { get; set; }

        public List<ConvertService.Result> Results { get; set; }

        public WindowPlacement WindowPlacement { get; set; }
    }

    // The rectangle uses desktop pixels, like the monitor bounds.
    internal class WindowPlacement
    {
        [JsonProperty(Required = Required.Always)]
        public double Left { get; set; }

        [JsonProperty(Required = Required.Always)]
        public double Top { get; set; }

        [JsonProperty(Required = Required.Always)]
        public double Width { get; set; }

        [JsonProperty(Required = Required.Always)]
        public double Height { get; set; }
    }
}
