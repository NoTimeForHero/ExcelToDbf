using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ExcelToDbf.Sources.Core.Data.Xml
{

    [XmlRoot("config", IsNullable = false)]
    public class Xml_Config
    {
        public bool log;
        public string LogLevel;

        public bool only_rules;
        public bool no_form_is_error;
        public bool show_messagebox_after;
        public bool skip_existing_files;

        public int buffer_size;

        public string title;
        public string status;
        public string warning;

        public Xml_OutFile outfile;

        [XmlArray]
        [XmlArrayItem("ext")]
        public List<string> extensions;

        [XmlArray]
        [XmlArrayItem("Form")]
        public List<Xml_Form> Forms;

        public static Xml_Config Load(String path)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Xml_Config));
            FileStream stream = new FileStream(path, FileMode.Open);
            Xml_Config container = serializer.Deserialize(stream) as Xml_Config;
            stream.Close();
            return container;
        }
    }

    public class Xml_OutFile
    {
        public bool simple;
        public string script;
    }

    public class Xml_Form
    {
        public string Name;

        [XmlArray]
        [XmlArrayItem("Equal", Type=typeof(Xml_Equal))]
        [XmlArrayItem("Group", Type=typeof(Xml_Equal_Group))]
        public List<Xml_Equal_Base> Rules;

        [XmlArray]
        [XmlArrayItem("field")]
        public List<Xml_DbfField> DBF;

        [XmlArray]
        [XmlArrayItem("Equal")]
        public List<Xml_Validator> Validate;

        public Xml_Form_Fields Fields;
    }

    public class Xml_Form_Fields
    {
        public Xml_Start_Y StartY;
        public int EndX;

        [XmlAnyElement("Static")]
        public XmlElement[] Static;

        [XmlAnyElement("Dynamic")]
        public XmlElement[] Dynamic;

        [XmlAnyElement("IF")]
        public XmlElement[] IF;
    }

    public class Xml_Start_Y : IXmlSerializable
    {
        public bool IsSimple => SimpleValue.HasValue;
        public int? SimpleValue;
        public Xml_Start_Y_Group group;

        public XmlSchema GetSchema() => null;

        public void ReadXml(XmlReader xmlReader)
        {
            var innerXML = xmlReader.ReadInnerXml();

            if (int.TryParse(innerXML, out int number))
            {
                SimpleValue = number;
                return;
            }

            using (TextReader textReader = new StringReader(innerXML))
            {
                XmlRootAttribute root = new XmlRootAttribute("Group");
                group = (Xml_Start_Y_Group) new XmlSerializer(typeof(Xml_Start_Y_Group), root).Deserialize(textReader);
            }
        }

        public void WriteXml(XmlWriter writer)
        {
            throw new NotImplementedException();
        }
    }

    public class Xml_Start_Y_Group
    {
        [XmlAttribute]
        public string name;

        [XmlAttribute]
        public string position;

        [XmlAttribute]
        public int Y;
    }

    public class Xml_Validator
    {
        [XmlAttribute]
        public string var1;

        [XmlAttribute]
        public string var2;

        public Xml_ValidatorMath Math;

        [XmlElement]
        public string Message;
    }

    public class Xml_ValidatorMath
    {
        [XmlAttribute]
        public int count;

        [XmlAttribute]
        public string precision;

        [XmlText]
        public string message;
    }

    public class Xml_DbfField
    {
        [XmlAttribute]
        public string name;

        [XmlAttribute]
        public string type;

        [XmlAttribute]
        public string length;

        [XmlAttribute]
        public string format;

        [XmlText]
        public string text;

        public Xml_DbfField()
        {
            format = "yyyy-MM-dd";
        }
    }

    public abstract class Xml_Equal_Base
    {
        [XmlIgnore]
        public int? X;

        [XmlIgnore]
        public int? Y;

        [XmlAttribute("X")]
        public string XmlSetterX
        {
            set => X = !string.IsNullOrEmpty(value) ? int.Parse(value) : default(int?);
            get => X?.ToString();
        }

        [XmlAttribute("Y")]
        public string XmlSetterY
        {
            set => Y = !string.IsNullOrEmpty(value) ? int.Parse(value) : default(int?);
            get => Y?.ToString();
        }
    }

    public class Xml_Equal_Group : Xml_Equal_Base
    {
        [XmlAttribute]
        public string Name;

        [XmlElement("Equal")]
        public List<Xml_Equal> Rules;
    }

    public class Xml_Equal : Xml_Equal_Base
    {
        [XmlText]
        public string Text;

        [XmlAttribute]
        public string validate;

        [XmlAttribute]
        public string regex_pattern;

        [XmlAttribute("regex_group")]
        protected string __string_regex_group;

        [XmlIgnore]
        protected int? __int_regex_group;

        [XmlIgnore]
        public int regex_group
        {
            get
            {
                if (!__int_regex_group.HasValue) __int_regex_group =
                        (__string_regex_group != null) ? int.Parse(__string_regex_group) : 1;
                return __int_regex_group.Value;
            }
        }
    }

}
