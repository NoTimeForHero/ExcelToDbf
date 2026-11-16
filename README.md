# Конвертер из Excel в DBF

Данная программа предназначена для экспорта данных из Excel форматов в DBF файлы для легаси приложений.  
Она использует Microsoft Office OLE для поддержки всех возможных Excel форматов (включая XML).

Для обработки документов нужно дополнить локальный `config.js`, создав форму с алгоритмом конвертации. У каждой фирмы может быть собственный конфиг.
При запуске процесса, программа пытается найти подходящую форму из списка для каждого документа в указанной директории.  
В `App/TestData` лежат примеры и `cases.json` с регрессионными ожиданиями. Они копируются в готовую поставку.

В Release-поставке GUI в папке `DLL` находится `ExcelToDbf.Cli.exe` для работы из PowerShell и LLM без исходников и рефлексии. Начните с [AGENTS.md](distribution/AGENTS.md), [справочника CLI](distribution/docs/cli.md), [JS API](distribution/docs/rules-api.md) и [описания DBF](distribution/docs/dbf-format.md).

```powershell
.\DLL\ExcelToDbf.Cli.exe doctor --json
.\DLL\ExcelToDbf.Cli.exe check-config --config .\config.js --json
.\DLL\ExcelToDbf.Cli.exe test --manifest .\TestData\cases.json --config .\config.example.js --json
```

После Release-сборки `App/bin/Release/App.zip` можно распаковать поверх директории приложения для обновления. Архив не содержит рабочий `config.js`, примерный конфиг или тестовые данные.

Для добавления новой формы используйте копию конфига и полный `validate`. Примерный конфиг обновляется как `config.example.js`; существующий рабочий `config.js` при сборке сохраняется. Миграция старых XML-конфигов выполняется отдельно.

Разработчикам: [сборка и проверки](docs/development.md), [план](docs/plans/2026-11-16-llm-friendly-cli.md), [результаты этапов](docs/plans/implementation-notes.md). Прежняя [XML-документация](DOCUMENTATION.md) оставлена как историческая.


## Возможные проблемы и их решение

Нужны Windows и установленный Microsoft Excel. Проекты GUI, ядра и CLI используют .NET Framework 4.7.2, вспомогательный CSV-конвертер — .NET Framework 3.5. Сборка и проверки выполнялись в Visual Studio 2022 Community с Excel 16.0; пути к Office interop задаются ссылками проектов.


## Используемые библиотеки
* [SocialExplorer/FastDBF](https://github.com/SocialExplorer/FastDBF) - работа с DBF файлами
* [Jint](https://github.com/sebastienros/jint) - JS интерпретатор правил распознавания, преобразования и выходного имени файла
* [NickBuhro.Translit](https://github.com/nick-buhro/Translit) - транслитерации согласно ГОСТу
* [Windows API Code Pack](https://github.com/aybe/Windows-API-Code-Pack-1.1) - диалога выбора директории из нового Windows API
* [Exception Assert](https://www.nuget.org/packages/ExceptionAssert/1.0.0) - библиотека, упрощающая юнит-тестирование


## Используемые иконки

Почти все иконки были взяты из [IconFinder](https://www.iconfinder.com/), с соблюдением лицензии.
Ссылки на авторов используемых иконок приведены в окне About.


## Лицензия

Данный проект распостраняется согласно лицензии [Apache 2.0](LICENSE.txt).
