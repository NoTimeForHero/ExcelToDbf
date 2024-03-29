﻿app.settings = {
    system: {
        outputEncoding: 866, // Номер выходная кодировка в DBF (число или название), по умолчанию 866 (1251 для Windows-1251, 65001 для UTF-8)
        bufferSize: 200, // Сколько записей будет обработано за один OLE запрос к Excel        
        extraWarning: 'Внимание! Не все файлы были сконвертированы!',
        noFormIsError: false, // Если true то вместо иконки Warning будет иконка Error
        fastSearch: false, // Если true, то проверка по списку правил будет прекращена после первого несоответствия, а по списку форм после первой найденной формы
    },
    logging: {
        enabled: true,
        level: 'tracer', // Уровень логгирования: TRACER|DEBUG|INFO|WARN|ERROR,
    },
    header: {
        title: 'ООО "Рога и копыта"',
        status: 'Пример информации'
    },
    extensions: [
        '*.xls',
        '*.xlsx',
        '*.xml'
    ]
}

// ====== БАЗОВЫЙ СПИСОК ФУНКЦИЙ =======
// string translit(input: string) - возвращает строку input в транслите
// string nospace(input: string, replaced: string) - заменяет в строке input все пробелы на replaced и возвращает строку
// string|null cell(y: int, x: int) - читает значение из ячейки Excel, возвращает null если произошла ошибка
// string|null matched(input: string, regex: Regex, id = 1: int) - разделяет строку input по регулярному выражению regex и возвращает id элемент полученного массива (1 если не указано) или null
// boolean match(input: string, regex: Regex) - попадает ли строка под регулярку
// boolean includes(input: string, search: string) - ищет подстроку search в строке input
// void log(message: string) - вывести сообщение через Console.WriteLine (по умолчанию)
// void error(message: string) - кидает исключение класса Jint.Runtime.JavaScriptException с сообщением message

app.getOutputFilename = function(file) {
        // Функция
        // На выход должна подаваться единственная строка с новым именем файла
        //
        // Функции (помимо базовых):
        // string|null dir(id: in) - возвращает сегмент пути по заданному индексу
        //
        // Переменные:
        // string file - оригинальное имя Excel файла
        // string dirCount - количество сегментов в пути
        file = nospace(file);
        file = translit(file);
        //file = "../" + file + ".dbf";
        file = file + ".dbf";
        return file;
}

app.forms = [
    {
        name: 'Форма 2.21А',
        settings: {
            startY: 8,
            endX: 7,
        },
        rules: function () {
            // Функции (помимо базовых):
            // Cell|null cell(y: int, x: int) - возвращает null при ошибки или interface Cell { x: int, y: int, value: object }
            // bool assert(current: string|Cell, expected: string, checkRegex: boolean = false)
            // Cell|null findRange(startY: int, startX: int, endY: int, endX: int) - поиск в указанном диапазоне для плавающего хедера

            // assert('Получилось', 'Ожидалось');
            assert(cell(2, 2), 'Форма 2.21А');
            assert(cell(7, 2), '№');
            assert(cell(7, 3), 'ФИО');
            assert(cell(7, 4), 'Счёт');
            assert(cell(7, 5), 'Сумма');
            assert(cell(7, 6), 'Дата оплаты');
        },
        dbfFields: [
            { name: 'ID', length: '8' },
            { name: 'KP', length: '8' },
            { name: 'FIO', length: '60' },
            { name: 'SUMMA', type: 'number', length: '10,2' },
            { name: 'DATEOPL', type: 'date' }
        ],
        // Функции (помимо базовых):
        // Cell|null cell(y: int, x: int) - возвращает Cell { x: int, y: int, value: object }
        // line: string[] - массив текущий XLS строки индексация стратует с 0
        // context: object - контекст доступный на протяжении всей обработки документа
        // stop() : null - остановка цикла записи
        write: function (line) {
            if (match(line[2], '^Данные от')) return null;
            if (includes(line[2], 'ИТОГО')) return stop();

            return {
                ID: line[2],
                FIO: line[3],
                KP: line[4],
                SUMMA: line[5],
                DATEOPL: line[6]
            }
        }
    },
    {
        name: 'Форма 2.21Б (с регулярками)',
        settings: {
            startY: 8,
            endX: 7,
        },
        rules: function () {
            assert(cell(2, 2), 'Форма 2.21Б');
            assert(cell(7, 2), '№');
            assert(cell(7, 3), 'ФИО');
            assert(cell(7, 4), 'Счёт');
            assert(cell(7, 5), 'Сумма');
            assert(cell(7, 6), 'Дата оплаты');
            assert(cell(6, 2), /на \d{2}\.\d{2}\.\d{4}/);
        },
        dbfFields: [
            { name: 'ID', length: '8' },
            { name: 'KP', length: '8' },
            { name: 'FIO', length: '60' },
            { name: 'SUMMA', type: 'number', length: '10,2' },
            { name: 'DATEOPL', type: 'date' },
            { name: 'NACHDATE', type: 'date' },
        ],
        beforeWrite: function() {
            context.sum = 0;
            context.NACHDATE = matches(cell(6, 2).Value, '\\d{2}\\.\\d{2}\\.\\d{4}')[0];
        },
        afterWrite: function () {
            // TODO: Упростить процесс суммирования?
            log("Сумма подсчитанная: " + context.sum);
            log("Сумма в документе: " + context.docSum);
            if (Math.abs(context.sum - context.docSum) > 5) error(`Подсчитанная сумма ${context.sum} отличается от ИТОГО ${context.docSum}!`);
        },
        write: function (line) {

            if (match(line[2], '^Данные от')) return null;
            if (!line[2]) return null;
            if (line[2] === 'Итого по дому:') return null;

            if (includes(line[2], 'ИТОГО')) {
                context.docSum = line[5];
                return stop();
            }

            if (line[2] === '###') {
                context.address = line[3];
                return null;
            }

            context.sum += parseFloat(line[5]) || 0;

            return {
                ID: line[2],
                FIO: line[3],
                KP: line[4],
                SUMMA: line[5],
                DATEOPL: line[6],
                NACHDATE: context.NACHDATE
            }
        }
    },
    {
        name: 'Форма 2.21В (плавающий заголовок)',
        settings: {
            startY: -1,
            endX: 7,
        },
        rules: function () {
            // Принудительно пропускаем чтобы не искать по ренджу заголовок если файл и так не подходит
            if (!assert(cell(2, 2), 'Форма 2.21В')) return;

            const header = findRange(4, 100, 2, 2, '№');
            if (!header) {
                assert('№', 'Заголовок не найден!');
                return;
            }
            const y = header.y;
            context.startY = y + 1;

            assert(cell(y, 2), '№');
            assert(cell(y, 3), 'ФИО');
            assert(cell(y, 4), 'Счёт');
            assert(cell(y, 5), 'Сумма');
            assert(cell(y, 6), 'Дата оплаты');
        },
        dbfFields: [
            { name: 'ID', length: '8' },
            { name: 'KP', length: '8' },
            { name: 'PERSON', length: '140' },
            { name: 'SUMMA', type: 'number', length: '10,2' },
            { name: 'DATEOPL', type: 'date' }
        ],
        write: function (line) {

            if (match(line[2], '^Данные от')) return null;
            if (!line[2]) return null;
            if (line[2] === 'Итого по дому:') return null;

            if (includes(line[2], 'ИТОГО')) return stop();

            if (line[2] === '###') {
                context.address = line[3];
                return null;
            }

            const PERSON = line[3] + ", живущая по адресу: " + context.address;

            return {
                ID: line[2],
                PERSON,
                KP: line[4],
                SUMMA: line[5],
                DATEOPL: line[6]
            }
        }
    },
]