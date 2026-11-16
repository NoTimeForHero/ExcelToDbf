using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToDbf.Processing
{
    public sealed class DbfFieldInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public int Length { get; set; }
        public int Decimals { get; set; }
    }

    public sealed class DbfInspection
    {
        public int Version { get; set; }
        public int LanguageDriver { get; set; }
        public int Encoding { get; set; }
        public int Records { get; set; }
        public List<DbfFieldInfo> Fields { get; } = new List<DbfFieldInfo>();
        public Dictionary<string, decimal> NumericTotals { get; } = new Dictionary<string, decimal>();
        public List<Dictionary<string, object>> Sample { get; } = new List<Dictionary<string, object>>();
    }

    // Independent read-back of the bytes written by FastDBF; no Excel or reflection.
    public static class DbfInspector
    {
        public static DbfInspection Read(string path, int codePage, int? expectedRecords = null)
        {
            using (var stream = File.OpenRead(path))
            using (var reader = new BinaryReader(stream))
            {
                var header = reader.ReadBytes(32);
                if (header.Length != 32 || header[0] != 3) throw Invalid("Ожидался заголовок dBASE III без memo");
                var count = BitConverter.ToInt32(header, 4);
                var headerLength = BitConverter.ToUInt16(header, 8);
                var recordLength = BitConverter.ToUInt16(header, 10);
                if (count < 0 || headerLength < 33 || (headerLength - 33) % 32 != 0 || recordLength < 1 ||
                    (long)headerLength + (long)recordLength * count > stream.Length)
                    throw Invalid("Повреждён размер DBF");
                if (expectedRecords.HasValue && expectedRecords != count) throw Invalid("Количество записей в DBF не совпадает с записанным");
                var inspection = new DbfInspection { Version = header[0], LanguageDriver = header[29], Encoding = codePage, Records = count };
                var encoding = Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                for (var offset = 32; offset < headerLength - 1; offset += 32)
                {
                    var descriptor = reader.ReadBytes(32);
                    var zero = Array.IndexOf(descriptor, (byte)0, 0, 11);
                    var field = new DbfFieldInfo { Name = Encoding.ASCII.GetString(descriptor, 0, zero < 0 ? 11 : zero),
                        Type = ((char)descriptor[11]).ToString(), Length = descriptor[16], Decimals = descriptor[17] };
                    if (field.Length == 0 || !new[] { "C", "N", "D" }.Contains(field.Type) ||
                        inspection.Fields.Any(x => x.Name == field.Name)) throw Invalid("Недопустимое описание поля DBF");
                    inspection.Fields.Add(field);
                    if (field.Type == "N") inspection.NumericTotals[field.Name] = 0;
                }
                if (reader.ReadByte() != 13 || inspection.Fields.Sum(x => x.Length) + 1 != recordLength)
                    throw Invalid("Повреждён заголовок полей DBF");
                for (var row = 0; row < count; row++)
                {
                    if (reader.ReadByte() != 32) throw Invalid("Неожиданная метка удаления записи DBF");
                    var sample = row < 5 ? new Dictionary<string, object>() : null;
                    foreach (var field in inspection.Fields)
                    {
                        var bytes = reader.ReadBytes(field.Length);
                        var text = encoding.GetString(bytes).TrimEnd(' ', '\0');
                        object value = text;
                        if (field.Type == "N")
                        {
                            text = text.Trim();
                            if (text.Length == 0) value = null;
                            else if (decimal.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                                CultureInfo.InvariantCulture, out var number))
                            { value = number; inspection.NumericTotals[field.Name] += number; }
                            else throw Invalid("Некорректное число после записи DBF: " + field.Name);
                        }
                        else if (field.Type == "D")
                        {
                            if (text.Trim().Length == 0) value = null;
                            else if (DateTime.TryParseExact(text, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                                value = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                            else throw Invalid("Некорректная дата после записи DBF: " + field.Name);
                        }
                        if (sample != null) sample[field.Name] = value;
                    }
                    if (sample != null) inspection.Sample.Add(sample);
                }
                return inspection;
            }
        }

        private static ProcessingException Invalid(string message) => new ProcessingException("INVALID_DBF", "dbf-readback", message);
    }
}
