using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;

namespace ExcelToDbf.Core.ViewModels
{
    internal class ConvertResultVM : ReactiveObject
    {
        [Reactive]
        public List<ConvertService.Result> Results { get; set; }

        [Reactive]
        public string Warning { get; set; }

        public ConvertResultVM()
        {
            Results = new List<ConvertService.Result>();
            var form = new DocForm { Name = "Форма 2.21А" };
            Results.Add(new ConvertService.Result
            {
                File = new FileModel { FileName = "Example Worksheet.xlsx" },
                OutputFilename = "Result.dbf",
                SearchResult = new SearchFormResult
                {
                    Result = form,
                    Report = new Dictionary<DocForm, List<SearchMatch>>
                    {
                        {form, new List<SearchMatch>
                        {
                            SearchMatch.Make("Строка 1", "Строка 2", true).With(1,1),
                            SearchMatch.Make("Строка 1", "Строка 2", true).With(2,1)
                        }}
                    }
                }
            });
            Warning = "Не все файлы были сконвертированы?!";
        }

        public ConvertResultVM(ConfigProvider cvConfig)
        {
            this.WhenAnyValue(x => x.Results).Subscribe((results) =>
            {
                var isAllConverted = results?.All(x => x.Error == null && x.Status == ConvertService.Result.ResultType.Converted) ?? true;
                Warning = isAllConverted ? null : cvConfig.Config.System.ExtraWarning;
            });
        }
    }
}
