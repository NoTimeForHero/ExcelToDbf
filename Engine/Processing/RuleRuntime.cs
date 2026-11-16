using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts.Context;
using Jint;
using NLog;
using Newtonsoft.Json;

namespace ExcelToDbf.Processing
{
    public sealed class RuleRuntime : IDisposable
    {
        public Engine Engine { get; }
        public ConfigContext Config { get; }
        public ExcelContext Context { get; }
        public string ConfigHash { get; }
        public string[] UnsupportedProperties { get; }

        public RuleRuntime(string path)
        {
            Engine = new Engine();
            try
            {
                var logger = LogManager.GetLogger("Rules");
                new GenericContext(Engine);
                Config = new ConfigContext(logger, Engine, path);
                Validate();
                UnsupportedProperties = JsonConvert.DeserializeObject<string[]>(Engine.Evaluate(@"
JSON.stringify((function() {
    var result = [];
    function check(value, keys, prefix) {
        Object.keys(value || {}).forEach(function(key) {
            if (keys.indexOf(key) < 0) result.push(prefix + '.' + key);
        });
    }
    check(app, ['settings', 'forms', 'getOutputFilename'], 'app');
    check(app.settings, ['system', 'header', 'extensions', 'logging'], 'app.settings');
    check(app.settings.system, ['outputEncoding', 'bufferSize', 'fastSearch', 'noFormIsError', 'extraWarning'], 'app.settings.system');
    check(app.settings.header, ['title', 'status'], 'app.settings.header');
    app.forms.forEach(function(form, index) {
        var prefix = 'app.forms[' + index + ']';
        check(form, ['name', 'settings', 'rules', 'dbfFields', 'beforeWrite', 'write', 'afterWrite'], prefix);
        check(form.settings, ['startY', 'endX'], prefix + '.settings');
        form.dbfFields.forEach(function(field, column) {
            check(field, ['name', 'type', 'length'], prefix + '.dbfFields[' + column + ']');
        });
    });
    return result;
})())").AsString());
                Context = new ExcelContext(logger, Config, Engine);
                using (var sha = SHA256.Create())
                    ConfigHash = BitConverter.ToString(sha.ComputeHash(File.ReadAllBytes(Config.ConfigPath))).Replace("-", "").ToLowerInvariant();
            }
            catch (Exception error)
            {
                Engine.Dispose();
                if (error is IOException || error is UnauthorizedAccessException) throw;
                throw new ProcessingException("INVALID_CONFIG", "config", error.Message, error);
            }
        }

        private void Validate()
        {
            if (Config.Data.Config?.System == null) throw new ArgumentException("app.settings.system обязателен");
            var system = Config.Data.Config.System;
            if (system.BufferSize < 1 || system.BufferSize > 100000)
                throw new ArgumentException("bufferSize должен быть в диапазоне 1..100000");
            Encoding.GetEncoding(system.OutputEncoding);
            if (Engine.Evaluate("typeof app.getOutputFilename").AsString() != "function")
                throw new ArgumentException("app.getOutputFilename должен быть функцией");
            if (Config.Forms.Length == 0) throw new ArgumentException("app.forms должен содержать хотя бы одну форму");
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var form in Config.Forms)
            {
                if (string.IsNullOrWhiteSpace(form.Name) || !names.Add(form.Name))
                    throw new ArgumentException("Имя формы обязательно и должно быть уникальным: " + form.Name);
                if (form.Rules == null || form.Write == null)
                    throw new ArgumentException(form.Name + ": rules и write должны быть функциями");
                if (form.Settings == null || form.Settings.EndX < 1 || form.Settings.EndX > 16384)
                    throw new ArgumentException(form.Name + ": settings.endX должен быть в диапазоне 1..16384");
                if (form.Settings.StartY == 0 || form.Settings.StartY < -1 || form.Settings.StartY > 1048576)
                    throw new ArgumentException(form.Name + ": startY должен быть -1 (из context) или номером строки Excel");
                if (form.Fields == null || form.Fields.Length == 0)
                    throw new ArgumentException(form.Name + ": dbfFields не должен быть пустым");
                var fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var field in form.Fields)
                {
                    if (field == null || !Regex.IsMatch(field.Name ?? "", "^[A-Za-z_][A-Za-z0-9_]{0,9}$") || !fields.Add(field.Name))
                        throw new ArgumentException(form.Name + ": недопустимое или повторное имя DBF-поля " + field?.Name);
                    var type = (field.Type ?? "string").ToLowerInvariant();
                    if (type != "string" && type != "number" && type != "date")
                        throw new ArgumentException(form.Name + ": неподдерживаемый тип поля " + type);
                    var length = field.Length ?? (type == "date" ? "8" : type == "number" ? "10,2" : "");
                    var parts = length.Split(',');
                    if (parts.Length > 2 || !int.TryParse(parts[0], out var width) || width < 1 || width > 254)
                        throw new ArgumentException(form.Name + ": недопустимая длина " + field.Name);
                    var decimals = 0;
                    if (parts.Length == 2 && !int.TryParse(parts[1], out decimals)) decimals = -1;
                    if (decimals < 0 || (type != "number" && parts.Length > 1) ||
                        (type == "date" && width != 8) ||
                        (type == "number" && (width > 20 || decimals > 15 || decimals > 0 && decimals + 2 > width)))
                        throw new ArgumentException(form.Name + ": недопустимые длина/точность поля " + field.Name);
                }
            }
        }

        public void Dispose() => Engine.Dispose();
    }
}
