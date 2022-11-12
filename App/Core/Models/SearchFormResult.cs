using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Utils.Serializers;
using Newtonsoft.Json;

namespace ExcelToDbf.Core.Models
{
    public class SearchFormResult
    {
        public DocForm Result { get; set; }

        [JsonConverter(typeof(DocFormDictionaryKeyConverter))]
        public Dictionary<DocForm, List<SearchMatch>> Report { get; set; } = new Dictionary<DocForm, List<SearchMatch>>();
    }
}
