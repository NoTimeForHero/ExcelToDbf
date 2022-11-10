using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDbf.Core.Models
{
    public class SearchFormResult
    {
        public FileModel Target { get; set; }
        public DocForm Result { get; set; }
        public Dictionary<DocForm, List<SearchMatch>> Report { get; set; } = new Dictionary<DocForm, List<SearchMatch>>();
    }
}
