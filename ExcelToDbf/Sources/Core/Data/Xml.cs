using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelToDbf.Sources.Core.Data
{

    [XmlRoot("config", IsNullable = false)]
    public class Xml_Config
    {
        public string inputDirectory;
        public string outputDirectory;

        public bool log;
        public string LogLevel;

        public bool only_rules;
        public bool save_memory;

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
        [XmlArrayItem("Equal")]
        public List<Xml_Equal> Rules;

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
        public int StartY;
        public int EndX;

        [XmlAnyElement("Static")]
        public XmlElement[] Static;

        [XmlAnyElement("Dynamic")]
        public XmlElement[] Dynamic;

        [XmlAnyElement("IF")]
        public XmlElement[] IF;
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

    public class Xml_Equal
    {
        [XmlAttribute]
        public int X;

        [XmlAttribute]
        public int Y;

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
