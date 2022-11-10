﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Jint;
using NickBuhro.Translit;
using NLog;
using NLog.Fluent;

namespace ExcelToDbf.Core.Services.Scripts.Context
{
    public class GenericContext : AbstractContext
    {

        public GenericContext(Engine engine) : base(engine)
        {
            engine.SetValue("translit", (Func<string, string>)FuncTranslit);
            engine.SetValue("nospace", (Func<string, string, string>)FuncReplaceSpace);
            engine.SetValue("afterRegEx", (Func<string, Regex, object, string>)FuncAfterRegEx);
            engine.SetValue("error", (Action<string>)FuncThrowException);
            AddLogger();
        }

        private void AddLogger()
        {
            var logger = LogManager.GetLogger(nameof(ScriptEngine));
            void Log(object data) => logger.Info($"{data}");
            engine.SetValue("log", (Action<string>)Log);
        }

        protected Regex regexSpace = new Regex(@"\s+", RegexOptions.Compiled);

        /// <summary>
        /// Переводит строку в транслит
        /// </summary>
        protected string FuncTranslit(string input)
        {
            return SafeString(Transliteration.CyrillicToLatin(input, Language.Russian));
        }

        /// <summary>
        /// Удаляет из строки все недопустимые для файловой системы символы
        /// </summary>
        protected string SafeString(string result)
        {
            Array.ForEach(Path.GetInvalidFileNameChars(),
                  c => result = result.Replace(c.ToString(), String.Empty));
            return result;
        }

        /// <summary>
        /// Заменяет все пробельные символы в строке на указанную строку
        /// </summary>
        protected string FuncReplaceSpace(string input, string replace) => regexSpace.Replace(input, replace ?? "");

        /// <summary>
        /// Разбивает подстроку input по регулярному выражению  info и возвращает nid группу
        /// Например: для построки abc с регуляркой one(two)(three)(four) и nid=2 вернёт "three"
        /// </summary>
        protected string FuncAfterRegEx(String input, Regex info, object nid)
        {
            int id = nid != null ? Convert.ToInt32(nid) : 1; // 1 == default
            string[] groups = info.Split(input);
            if (id > groups.Length - 1) return null;
            return groups[id];
        }

        /// <summary>
        /// Бросает исключение с заданным сообщением
        /// </summary>
        protected void FuncThrowException(String text)
        {
            throw new JSException("Исключение вызванное из JavaScript: " + text);
        }
    }
}
