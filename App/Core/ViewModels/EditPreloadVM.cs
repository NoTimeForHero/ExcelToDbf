using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Services;
using NLog;
using ReactiveUI;
using ReactiveUI.Fody.Helpers;
using PConfig = ExcelToDbf.Core.Services.Preload.Config;
using Repository = ExcelToDbf.Core.Services.Preload.Repository;

namespace ExcelToDbf.Core.ViewModels
{
    public class EditPreloadVM : ReactiveObject
    {
        private readonly PreloadService srvPreload;

        [Reactive]
        public PConfig Config { get; set; }

        [Reactive]
        public Repository VRepository { get; set; }

        [Reactive]
        public bool IsLoading { get; set; }

        [Reactive]
        public bool RepositoryDirty { get; set; }

        [Reactive]
        public Repository.Tag SelectedTag { get; set; }

        [Reactive]
        public Dictionary<string,string> AvailableVersions { get; set; }

        [Reactive]
        public KeyValuePair<string,string>? SelectedVersion { get; set; }

        [Reactive]
        public string Error { get; set; } = "";

        public ReactiveCommand<Unit, Unit> ReloadWithVersionCommand { get; set; }
        public ReactiveCommand<Unit, Unit> ReloadCommand { get; set; }
        public ReactiveCommand<Unit, Unit> LoadRepositoryCommand { get; set; }

        public EditPreloadVM()
        {
            Config = new PConfig
            {
                Enabled = true,
                UseForceURL = false,
                ForceURL = "http://example.org/repository/firm1/latest.js",
                Repository = "http://example.org/repository/index.json",
                Tag = "Рога и копыта",
                Version = "latest"
            };
            VRepository = new Repository
            {
                Title = "Базовый репозиторий ООО \"Рога и копыта\"",
                Description = "Тут содержатся все основные формы организации",
                Root = "http://example.org/repository/firm1/",
                Tags = new List<Repository.Tag>
                {
                    new Repository.Tag {
                        Title = "Бухгалтерия",
                        Url = "accounting",
                        Versions = new Dictionary<string, string>
                        {
                            {"latest", "latest.js"},
                            {"beta", "version_231122.js"}
                        }
                    },
                    new Repository.Tag
                    {
                        Title = "Мененджмент",
                        Url = "managers",
                        Versions = new Dictionary<string, string>
                        {
                            {"latest", "latest.js"},
                        }
                    }
                }
            };
            SelectedTag = VRepository.Tags[0];
            SelectedVersion = new KeyValuePair<string, string>("beta", "version_231122.js");
        }

        private async Task FetchRepository()
        {
            try
            {
                Error = "";
                IsLoading = true;
                VRepository = await srvPreload.GetRepository(Config.Repository);
                RepositoryDirty = false;
            }
            catch (Exception ex)
            {
                // logger.Info("Ошибка проверки URL: " + ex.Message);
                Error = ex.Message;
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task Initialize()
        {
            Config = srvPreload.Settings;
            LoadRepositoryCommand = ReactiveCommand.CreateFromTask(FetchRepository);
            ReloadCommand = ReactiveCommand.CreateFromTask(() => srvPreload.Run());

            var canLoadVersion = this.WhenAnyValue(
                x => x.SelectedTag,
                x => x.SelectedVersion,
                (tag, version) => tag != null && version != null
            );
            ReloadWithVersionCommand = ReactiveCommand.CreateFromTask(() => srvPreload.Run(), canExecute: canLoadVersion);

            this.WhenAnyValue(x => x.Config.Repository)
                .Skip(1)
                .Subscribe(x =>
                {
                    SelectedTag = null;
                    SelectedVersion = null;
                    RepositoryDirty = true;
                });

            this.WhenAnyValue(x => x.SelectedTag)
                .Skip(1)
                .Subscribe(tag =>
                {
                    AvailableVersions = tag?.Versions;
                    Config.Tag = tag?.Title;
                });

            this.WhenAnyValue(x => x.SelectedVersion)
                .Skip(1)
                .Subscribe(pair => Config.Version = pair?.Key);

            RepositoryDirty = true;
            await FetchRepository();
            SelectedTag = VRepository?.Tags.FirstOrDefault(x => x.Title == Config.Tag);
            SelectedVersion = AvailableVersions?.ContainsKey(Config.Version) ?? false
                ? new KeyValuePair<string,string>(Config.Version, AvailableVersions[Config.Version])
                : (KeyValuePair<string, string>?)null;
        }

        public EditPreloadVM(PreloadService srvPreload)
        {
            this.srvPreload = srvPreload;
            var _ = Initialize();
        }
    }
}
