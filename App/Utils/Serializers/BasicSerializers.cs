using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelToDbf.Utils.Serializers
{
    public class DocFormConverter : JsonConverter<DocForm>
    {
        public override void WriteJson(JsonWriter writer, DocForm value, JsonSerializer serializer)
        {
            writer.WriteStartObject();
            writer.WritePropertyName("Name");
            writer.WriteValue(value.Name);
            writer.WriteEndObject();
        }

        public override DocForm ReadJson(JsonReader reader, Type objectType, DocForm existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null) return null;
            if (reader.TokenType != JsonToken.StartObject) throw new InvalidOperationException("DocForm must be object!");
            var data = JObject.Load(reader);
            return new DocForm
            {
                Name = data["Name"]?.ToObject<string>()
            };
        }
    }

    public class DocFormDictionaryKeyConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object input, JsonSerializer serializer)
        {
            writer.WriteStartObject();
            foreach (DictionaryEntry item in (IDictionary)input)
            {
                writer.WritePropertyName(JsonConvert.SerializeObject(item.Key));
                serializer.Serialize(writer, item.Value);
            }
            writer.WriteEndObject();
        }

        // https://stackoverflow.com/a/38265583
        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            var types = objectType.GetGenericArguments();
            if (types.Length < 2) throw new InvalidOperationException("Dictionary<T,T2> must have 2 types!");
            var keyType = types[0];
            var valueType = types[1];
            var dictionaryType = typeof(Dictionary<,>).MakeGenericType(keyType, valueType);
            var instance = (IDictionary) Activator.CreateInstance(dictionaryType);
            foreach (var pair in JObject.Load(reader))
            {
                var key = JsonConvert.DeserializeObject(pair.Key, keyType);
                var value = pair.Value?.ToObject(valueType);
                instance.Add(key, value);
            }
            return instance;
        }

        public override bool CanConvert(Type objectType)
        {
            throw new NotImplementedException();
        }
    }
}
