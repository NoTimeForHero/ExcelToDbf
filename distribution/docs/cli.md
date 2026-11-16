# CLI конвертера

Среда: Windows, установленный Microsoft Excel и .NET Framework 4.7.2. CSV дополнительно использует поставляемый `CSV_Converter.exe` на .NET Framework 3.5. `doctor` проверяет реальное создание и закрытие Excel; успешный `check-config` Excel не требует.

Команды ниже выполняются из корня Release-поставки GUI. В Debug-поставке и самостоятельной сборке CLI EXE находится в корне: используй `.\ExcelToDbf.Cli.exe`.

## Быстрый запуск из PowerShell

```powershell
.\DLL\ExcelToDbf.Cli.exe doctor --json
New-Item -ItemType Directory -Path .\work -Force
Copy-Item -LiteralPath .\config.js -Destination .\work\candidate.js
.\DLL\ExcelToDbf.Cli.exe inspect .\incoming\new.xlsx --range A1:M30 --json
# Отредактировать work\candidate.js по rules-api.md.
.\DLL\ExcelToDbf.Cli.exe check-config --config .\work\candidate.js --json
.\DLL\ExcelToDbf.Cli.exe match .\incoming\new.xlsx --config .\work\candidate.js --json
.\DLL\ExcelToDbf.Cli.exe validate .\incoming\new.xlsx --config .\work\candidate.js --json
.\DLL\ExcelToDbf.Cli.exe test --manifest .\TestData\cases.json --config .\work\candidate.js --json
.\DLL\ExcelToDbf.Cli.exe convert .\incoming\new.xlsx --config .\work\candidate.js --output .\work\result.dbf --json
```

Manifest примера подходит только к примерному конфигу. Для фирмы используй её набор данных. После каждой команды проверяй `$LASTEXITCODE`; следующий шаг выполняй после разбора предыдущего результата.

Для запуска эталонных сценариев поставки:

```powershell
.\DLL\ExcelToDbf.Cli.exe test --manifest .\TestData\cases.json --config .\config.example.js --json
```

## Общие аргументы и пути

CLI пишет JSON в UTF-8. Для разбора JSON в Windows PowerShell 5.1 сначала настрой кодировку; PowerShell 7 обычно уже использует UTF-8:

```powershell
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding
```

Windows PowerShell 5.1 представляет native stderr как ErrorRecord. Если в скрипте задано `$ErrorActionPreference = 'Stop'`, сообщения логов могут прервать вызов даже при коде 0. На время вызова EXE используй `Continue`, сохраняй stderr отдельно и проверяй `$LASTEXITCODE`; затем восстанови прежнюю настройку. Это особенность оболочки, сообщения логов не являются ошибкой конвертации.

Сам CLI пишет JSON в UTF-8, но обёртка на Python или Node.js может печатать разобранный результат в текущей кодовой странице консоли — тогда русский текст в выводе будет испорчен при корректных данных. Для Python задай `PYTHONIOENCODING=utf-8`, для других сред — `chcp 65001` или явную кодировку вывода.

- `--help` работает у любой команды без выполнения операции.
- `--json` оставляет в stdout ровно один JSON-документ. Логи, сообщения `log()` и название выполняемого теста идут в stderr. Потоки не объединять перед разбором JSON.
- `--timeout N` задаёт предел операции в секундах, от 1 до 86400; по умолчанию 120. Для `test` это предел всего набора. По тайм-ауту рабочий процесс и зарегистрированные принадлежащие операции процессы останавливаются.
- `--config path` выбирает локальный JS-конфиг; поддерживается во всех рабочих командах, кроме `inspect`. Без аргумента используется `config.js` в корне поставки (для самостоятельной сборки — рядом с EXE).
- Явные относительные пути считаются от рабочей директории вызывающего процесса. Путь конфига по умолчанию не зависит от неё. В JSON возвращаются абсолютные пути.
- CLI работает локально, не запускает GUI, предзагрузку конфигурации и автообновление. Запуск не заменяет рабочий конфиг.
- Сборка сохраняет существующий локальный `config.js` и отдельно обновляет `config.example.js`. При ручном обновлении поставки также сохраняй конфиги фирм и предыдущие версии.

## Команды

| Команда | Аргументы | Результат |
| --- | --- | --- |
| `doctor` | необязательный `--config` | Доступность Excel и временной директории, версия Excel, разрядность CLI, пути |
| `check-config` | `--config` | Синтаксис и структура JS, параметры чтения, схемы DBF, список форм, SHA256 |
| `inspect file` | `--range A1:M30` | Листы, выбранный первый лист, UsedRange, координаты и значения ячеек |
| `match file` | `--config`, `--all-forms`, `--strict` | Выбор формы и объяснение assertions |
| `validate file` | `--config`, `--strict`, `--keep-artifacts dir` | Полная пробная конвертация и чтение DBF обратно |
| `convert file` | как validate, также `--output path`, `--overwrite` | Проверенная конвертация с публикацией DBF |
| `test` | `--manifest path`, `--config`, `--strict` | Сравнение каждого документа с заданными ожиданиями |

У `inspect` по умолчанию диапазон A1:M30; максимум 10000 ячеек за запрос. Поддержаны A1-адреса прямоугольника и одной ячейки. `rawValue` содержит значение Excel `Range.Value`, `ruleValue` — строку, которую вернёт `cell().Value`. `state` различает пустую ячейку, значение и ошибку Excel; ошибка чтения диапазона возвращается в `errors`. Отображаемый текст и формулы отдельно не экспортируются. Выбор другого листа первой версией не поддерживается.

Возвращается весь запрошенный прямоугольник, включая пустые ячейки и ячейки за пределами `usedRange` (`state: empty`). Поэтому при подсчёте заполненных строк и столбцов пересекай список `cells` с `usedRange`, иначе пустые строки посчитаются дважды. Поле `type` (`String`, `Double`, `DateTime`) показывает, чем ячейка является в Excel: по нему видно, дата в документе текстовая или типизированная, и какое преобразование получит значение в `write`.

`match` воспроизводит рабочий поиск с настройкой `fastSearch`. В `search.forms` только проверенные формы; `notExecutedForms` перечисляет остальные. У каждой проверенной формы есть выполненные assertions и признак `shortCircuited`. Досрочный `return` внутри JS не позволяет узнать, сколько assertions могло быть выполнено дальше. `--all-forms` добавляет `diagnosticSearch` из отдельного JS runtime с отключённым fastSearch, сохраняя рабочий `search`.

`validate` проходит весь документ и выполняет оба хука. DBF пишется во временную директорию, а не в рассчитанный `plannedOutput`. После закрытия файл проверяется отдельным читателем: заголовок, схема, все записи, числа, даты и кодировка. В отчёте есть количество строк, записей, пропусков, причина остановки, сумма каждого числового поля и первые пять записей. Ограничение образца не ограничивает проверку.

`conversion.rowsRead` — строки от `startY` до конца используемого диапазона. `conversion.rowsSkipped` — все пропущенные строки: пустые строки-разделители, повторные шапки, итоговые строки и подписи. `conversion.recordsWritten` равно `rowsRead - rowsSkipped`. На многостраничном реестре пропусков бывает больше, чем строк в шапке, поэтому ожидания manifest строй по фактическим числам, а не по числу страниц.

`--strict` превращает предупреждения о неоднозначном распознавании и потерях данных в ошибку. В обычном режиме в JSON сохраняется до 200 предупреждений; `dataWarningCount` сообщает полное число предупреждений записи. `--keep-artifacts dir` сохраняет отдельный DBF с `artifact.status=validated` или `partial-failed`. Сбойный артефакт не является готовым файлом для CRM. После принудительного тайм-аута сохранение артефакта не гарантируется.

`convert` без `--output` использует `app.getOutputFilename`. Относительный результат считается от директории входного файла, поэтому возвращённое `../name.dbf` для `incoming\2021\07-Июль\doc.xlsx` даёт `incoming\2021\name.dbf` — файл появится вне каталога входа. Сверяй полный путь в `plannedOutput` у `validate`; для испытаний задавай `--output` внутри тестового каталога. Выходная директория должна существовать. Совпадение выхода с исходником или конфигом запрещено. Для существующего файла нужен `--overwrite`; замена выполняется после успешной записи и `afterWrite` через временный файл в той же директории. При сбое прежний DBF сохраняется. Политика перезаписи GUI остаётся прежней.

## JSON и коды завершения

Каждый отчёт содержит `schemaVersion: 1`, `rulesApiVersion: 1`, `command`, `status`, `data`, `warnings`, `errors`. Отсутствующие необязательные поля не выводятся. Имена DBF-полей в словарях сохраняют регистр, например `numericTotals.SUMMA`.

| Код процесса | Значение |
| --- | --- |
| 0 | Успешное выполнение / все ожидания manifest соблюдены |
| 1 | Конфиг, правила или данные не прошли проверку, форма не найдена, регрессия обнаружила отличие |
| 2 | Ошибка аргументов CLI |
| 3 | Проблема среды, файлов, COM, тайм-аут или отмена |

Основные диагностические коды: `INVALID_ARGUMENTS`, `INVALID_CONFIG`, `UNSUPPORTED_PROPERTY`, `INVALID_RANGE`, `RULES_FAILED`, `NO_FORM`, `NO_ASSERTIONS`, `MULTIPLE_FORMS`, `INVALID_START_ROW`, `INVALID_NUMBER`, `NUMBER_OVERFLOW`, `DBF_FIELD_ERROR`, `INVALID_DATE`, `STRING_TRUNCATED`, `ENCODING_LOSS`, `NUMBER_PRECISION_LOSS`, `CONVERSION_FAILED`, `INVALID_DBF`, `OUTPUT_EXISTS`, `OUTPUT_NAME_FAILED`, `INVALID_OUTPUT`, `STRICT_VALIDATION`, `INVALID_MANIFEST`, `REGRESSION_FAILED`, `TIMEOUT`, `CANCELLED`, `EXECUTION_FAILED`, `ARTIFACT_SAVE_FAILED`, `CLEANUP_FAILED`.

Ошибка по возможности содержит `stage`, `form`, `sourceRow`, `field`, `actual`. `sourceRow` — номер строки исходного Excel. Координаты JS, когда доступны, находятся в `scriptFile`, `scriptLine`, `scriptColumn`, `scriptStack`; строки и колонки в отчёте считаются с 1. Не у каждого исключения хост-функции есть JS-позиция.

Основные значения `stage`: `rules`, `beforeWrite`, `write`, `dbf-write`, `afterWrite`, `output-name`. Сообщение `error(...)` из `rules` даёт `RULES_FAILED`, из хуков и `write` — `CONVERSION_FAILED` с соответствующим `stage`, из `app.getOutputFilename` — `OUTPUT_NAME_FAILED`.

Пример ошибки:

```json
{
  "code": "INVALID_NUMBER",
  "stage": "dbf-write",
  "form": "Форма 2.21А",
  "sourceRow": 9,
  "field": "SUMMA",
  "actual": "ERROR",
  "message": "На строке 9 колонка SUMMA (ERROR) не является числом"
}
```

## Структура `data`

Состав `data` стабилен для версии 1, но необязательные ключи могут отсутствовать. Для `validate` и `convert`:

| Ключ | Содержимое |
| --- | --- |
| `input`, `config`, `configHash`, `selectedSheet` | абсолютные пути, SHA256 конфига и выбранный лист |
| `search` | `selectedForm`, `matchedForms`, `fastSearch`, `forms[]`, `notExecutedForms`; у формы — `matched`, `shortCircuited` и `assertions[]` с `y`, `x`, `expected`, `actual`, `matches` |
| `plannedOutput` | путь из `app.getOutputFilename`; только расчёт, файл не создаётся |
| `dataWarningCount` | полное число предупреждений по данным (список в `warnings` ограничен 200) |
| `conversion` | `form`, `startY`, `rowsRead`, `rowsSkipped`, `recordsWritten`, `stopReason`, `sample` |
| `dbf` | `version`, `languageDriver`, `encoding`, `records`, `fields[]` (`name`, `type`, `length`, `decimals`), `numericTotals`, `sample` |

`stopReason` равен `stop`, если строки завершил `stop()`, и `end-of-sheet`, если обработка дошла до конца используемого диапазона. `sample` ограничен первыми пятью записями: `conversion.sample` — значения, возвращённые `write`, `dbf.sample` — значения прочитанного обратно файла (обрезанные строки, даты `yyyy-MM-dd`, числа как JSON number). У `test` в `data.cases[]` появляются `passed`, `differences[]` (`path`, `expected`, `actual`), `actualStatus`, `actualExitCode` и `actualData` с полным отчётом по документу.

## Manifest регрессионных проверок

Manifest — JSON с `schemaVersion: 1` и непустым массивом `cases`. Имена сценариев уникальны, `input` считается относительно manifest. Все сценарии выполняются с выбранным конфигом фирмы. Неизвестные свойства отклоняются, чтобы опечатка не отключила проверку.

```json
{
  "schemaVersion": 1,
  "cases": [
    {
      "name": "Новый формат",
      "input": "new.xlsx",
      "expected": {
        "status": "ok",
        "form": "Новая форма",
        "records": 10,
        "startY": 8,
        "stopReason": "stop",
        "warningCount": 0,
        "tolerance": 0,
        "numericTotals": { "SUMMA": 1234.56 },
        "samples": [ { "record": 1, "values": { "ID": "1", "SUMMA": 100.00 } } ]
      }
    },
    {
      "name": "Неподходящий файл",
      "input": "other.xlsx",
      "expected": { "status": "failed", "error": { "code": "NO_FORM", "stage": "rules" } }
    }
  ]
}
```

Для `ok` обязательны `form` и `records`; остальные ожидания необязательны. Для `failed` обязателен `error.code`; можно уточнить `stage`, `sourceRow`, `field`. `samples.record` — номер записи DBF от 1 до 5. Числа сравниваются с `tolerance`, по умолчанию точно. В `data.cases[].differences` перечислены несовпавшие ожидания. Общий успех означает, что ожидаемые ошибки тоже подтверждены.

`samples` сравниваются со значениями, **прочитанными из записанного DBF**, а не с тем, что вернул `write`. Поэтому в ожиданиях строка уже обрезана до длины поля, дата записана как `yyyy-MM-dd`, число — как JSON number. Эталонные образцы бери из `dbf.sample` у `validate` или `convert`: `conversion.sample` содержит исходные значения и на обрезке не совпадёт. Формат расхождения: `{ "path": "sample[1].ADRES", "expected": "...", "actual": "..." }`, номер записи в `path` считается от 1. Задавай `warningCount`: без него изменение числа потерь (например, обрезки) регрессия не заметит.

Проверки поставки фиксируют поведение примеров проекта и не подтверждают приём DBF конкретной коммерческой CRM. Для такого подтверждения нужен проверенный бизнес-эталон или импорт в CRM.
