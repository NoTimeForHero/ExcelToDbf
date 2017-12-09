using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using ExcelToDbf.Sources.Core;
using ExcelToDbf.Sources.Core.Data.Xml;
using ExcelToDbf.Sources.Core.External;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests
{
    [TestClass]
    public class UnitTest1
    {
        protected Xml_Form generateForm()
        {
            Xml_Form form = new Xml_Form
            {
                DBF = new List<Xml_DbfField>
                {
                    new Xml_DbfField
                    {
                        type = "string",
                        name = "FIO",
                        length = "30",
                        text = "$FIO"
                    },
                    new Xml_DbfField
                    {
                        type = "string",
                        name = "KP",
                        length = "8",
                        text = "$KP"
                    },
                    new Xml_DbfField
                    {
                        type = "numeric",
                        name = "SUMMA",
                        length = "8",
                        text = "$SUMMA"
                    },
                    new Xml_DbfField
                    {
                        type = "date",
                        name = "DATA",
                        length = "8",
                        text = "$DATA"
                    },
                    new Xml_DbfField
                    {
                        type = "date",
                        name = "NACHIS",
                        length = "8",
                        text = "$NACHIS",
                    }
                },

                // TODO: Переписать на нормальные объекты
                Fields = new Xml_Form_Fields
                {
                    EndX = 8,
                    StartY = 8,
                    Static = new[]
                    {
                        GetElement(@"<Static X=""2"" Y=""6"" name=""NACHIS"" type=""date"" regex_pattern=""на (\S+)"" language=""ru-RU"" format=""dd.MM.yyyy"" lastday=""true"" />")
                    },
                    IF = new[]
                    {
                        GetElement(@"<IF X=""2"" VALUE=""Итого:""><THEN><STOP_LOOP/><Dynamic X=""5"" name=""XLS_SUMMA"" type=""numeric"" /></THEN></IF>"),
                        GetElement(@"<IF X=""2"" VALUE=""Пропуск"">
                                        <THEN><SKIP_RECORD/></THEN>
                                        <ELSE><Dynamic X=""2"" name=""ID"" type=""string"" />
                                      </ELSE></IF>")
                    },
                    Dynamic = new[]
                    {
                        GetElement(@"<Dynamic X=""3"" name=""FIO"" type=""string"" />"),
                        GetElement(@"<Dynamic X=""4"" name=""KP"" type=""string"" />"),
                        GetElement(@"<Dynamic X=""5"" name=""TOTAL_SUMMA"" type=""numeric"" function=""SUM"" />"),
                        GetElement(@"<Dynamic X=""5"" name=""SUMMA"" type=""numeric"" />"),
                        GetElement(@"<Dynamic X=""6"" name=""DATA"" type=""date"" />"),
                    }
                },

                Validate = new List<Xml_Validator>
                {
                    new Xml_Validator
                    {
                        Math = new Xml_ValidatorMath
                        {
                            count  = 500,
                            precision = "0,5",
                            message = "Погрешность {1} при допуске {0}"
                        },
                        Message = "{0} != {1} !",
                        var1 = "TOTAL_SUMMA",
                        var2 = "XLS_SUMMA"
                    }
                }
            };
            return form;
        }

        protected XmlElement GetElement(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            return doc.DocumentElement;
        }


        [TestMethod]
        public void TestMethod1()
        {
            //Logger.SetFile(null);
            string pathExcel = Path.Combine(Environment.CurrentDirectory, "Data\\Example1.xlsx");
            string pathTemp = TestLibrary.getTempFilename(".dbf");

            Xml_Form form = generateForm();
            Excel excel = new Excel(true);
            DBF dbf = new DBF(pathTemp, form.DBF);
            try
            {
                excel.OpenWorksheet(pathExcel);

                Work work = new Work(form, 120);
                Logger.LogLevel old_level = Logger.Level;
                Logger.SetLevel(Logger.LogLevel.TRACER);
                work.IterateRecords(excel.worksheet, dbf.appendRecord);
                Logger.SetLevel(old_level);
            }
            finally
            {
                dbf.close();
                excel.close();
            }
        }
    }
}
