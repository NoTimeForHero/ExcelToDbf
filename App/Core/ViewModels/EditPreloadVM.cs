using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;
using PConfig = ExcelToDbf.Core.Services.Preload.Config;

namespace ExcelToDbf.Core.ViewModels
{
    public class EditPreloadVM : ReactiveObject
    {
        [Reactive]
        public PConfig Config { get; set; }

        [Reactive]
        public bool IsLoading { get; set; }

        public EditPreloadVM()
        {
            Config = new PConfig
            {
                Enabled = true,
                ForceURL = "http://example.org/repository/firm1/latest.js",
                Repository = "http://example.org/repository/index.json",
                Tag = "Рога и копыта",
                Version = "latest"
            };
        }

        public EditPreloadVM(PreloadService service)
        {
            Config = service.Settings;
        }
    }
}
