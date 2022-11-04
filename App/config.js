﻿app.settings = {
    system: {
        outputEncoding: 866, // Номер выходная кодировка в DBF (число или название), по умолчанию 866 (1251 для Windows-1251, 65001 для UTF-8)
        bufferSize: 2000, // Сколько записей будет обработано за один OLE запрос к Excel        
        extraWarning: 'Ошибка?',
        noFormIsError: false, // Если true, значит при пропуске хоть одного документа из-за отсутствия форм будет показана ошибка из элемента выше
    },
    logging: {
        enabled: true,
        level: 'tracer', // Уровень логгирования: TRACER|DEBUG|INFO|WARN|ERROR,
    },
    header: {
        title: 'ООО "Рога и копыта" 222222',
        status: 'Версия 0.0.0.1 Альфа\nПоследнее обновление 01.01.2020\nИнформация...'
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
        file = "../" + file + ".dbf";
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
            // void assert(current: string, expected: string, checkRegex: boolean = false)
            // void assertCell(y: int, x: int, expected: string, checkRegex: boolean = false)
            assertCell(2, 2, 'Форма 2.21А');
            assertCell(7, 2, '№');
            assertCell(7, 3, 'ФИО');
            assertCell(7, 4, 'Счёт');
            assertCell(7, 5, 'Сумма');
            assertCell(7, 6, 'Дата оплаты');
        },
        dbfFields: [
            {name: 'ID', length: '8' },
            {name: 'KP', length: '8' },
            {name: 'FIO', length: '60' },
            {name: 'SUMMA', type: 'numeric', length: '10,2' },                
            {name: 'DATEOPL', type: 'date' },        
        ],
        // Функции (помимо базовых):
        // line: string[] - массив текущий XLS строки индексация стратует с 1
        // context: unknown - резерв для будуших целей
        // stop() : null - остановка цикла записи
        write: function(line, context, stop) {
            if (match(line[2], '^Данные от')) return null;
            // TODO: Проверка суммы на совпадения с XLS файлом
            if (includes(line[2], 'ИТОГО')) return stop();

            return {
                ID: line[2],
                FIO: line[3],
                KP: line[4],
                SUMMA: line[5],
                DATEOPL: line[6],
            }
        }
    }
]