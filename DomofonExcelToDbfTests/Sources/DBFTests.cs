using Microsoft.VisualStudio.TestTools.UnitTesting;
using DomofonExcelToDbf.Sources;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DomofonExcelToDbf.Sources.Core;
using SocialExplorer.IO.FastDBF;
using DomofonExcelToDbf.Sources.Xml;

namespace DomofonExcelToDbf.Sources.Tests
{
    public class TestRepository
    {

        public static List<Xml_DbfField> getFields()
        {
            List<Xml_DbfField> data = new List<Xml_DbfField>
            {
                new Xml_DbfField { name = "fio",   type = "string",  length = "40", text = "$FIO"},
                new Xml_DbfField { name = "summa", type = "numeric", length = "10,4", text = "$SUMMA"},
                new Xml_DbfField { name = "data",  type = "date", length = "8", text = "$DATE"}
            };
            return data;
        }

        public static Dictionary<string, TVariable> getVariables()
        {
            Dictionary<string, TVariable> data = new Dictionary<string, TVariable>();

            var tvariable = new TVariable("FIO");
            tvariable.Set("Ivanov Ivan Ivanovich");
            data.Add(tvariable.name, tvariable);

            var tnumeric = new TNumeric("SUMMA");
            tnumeric.Set(12.3456f);
            data.Add(tnumeric.name, tnumeric);

            var tdate = new TDate("DATE");
            tdate.Set("22.11.2001");
            data.Add(tdate.name, tdate);

            return data;
        }
    }

    [TestClass]
    public class DBFWritingTests
    {
        private string dbfFileName;
        private DBF dbf;
        private Encoding encoding;
        private List<Xml_DbfField> fields;
        private Dictionary<string, TVariable> variables;

        [TestInitialize()]
        public void Startup()
        {
            Logger.instance = new Logger(level: Logger.LogLevel.DEBUG);

            encoding = Encoding.UTF8;
            fields = TestRepository.getFields();
            variables = TestRepository.getVariables();
            dbfFileName = Path.GetTempFileName();

            dbf = new DBF(dbfFileName, fields, encoding);
            dbf.appendRecord(variables);
            dbf.close();
        }

        [TestMethod]
        public void IsHeadersCorrect()
        {
            DbfFile dbfFile = new DbfFile(encoding);
            dbfFile.Open(dbfFileName, FileMode.Open);

            Assert.AreEqual(fields.Count, dbfFile.Header.ColumnCount);

            Assert.AreEqual(dbfFile.Header[0].Name, fields[0].name);
            Assert.AreEqual(dbfFile.Header[1].Name, fields[1].name);
            Assert.AreEqual(dbfFile.Header[2].Name, fields[2].name);
        }

        [TestMethod]
        public void IsDataCorrect()
        {
            DbfFile dbfFile = new DbfFile(encoding);
            dbfFile.Open(dbfFileName, FileMode.Open);
            DbfRecord orec = new DbfRecord(dbfFile.Header);

            Assert.IsTrue(dbfFile.ReadNext(orec));


            Assert.AreEqual(DbfColumn.DbfColumnType.Character, orec.Column(0).ColumnType);
            Assert.AreEqual(DbfColumn.DbfColumnType.Number, orec.Column(1).ColumnType);
            Assert.AreEqual(DbfColumn.DbfColumnType.Date, orec.Column(2).ColumnType);

            // DBF возвращает строки такой длины, какая указана в хедерах при создании 
            string fio = "Ivanov Ivan Ivanovich";
            Assert.AreEqual(fio + new String(' ', 40 - fio.Length), orec[0]);

            string num = "12.3456";
            Assert.AreEqual(new String(' ',10 - num.Length) + num, orec[1]);

            Assert.AreEqual("20011122", orec[2]);
        }

        [TestMethod]
        public void RepeatClose()
        {
            dbf.close();
            dbf.close();
        }

        [TestMethod]
        public void IsFileCreated()
        {
            bool exists = File.Exists(dbfFileName);
            Assert.IsTrue(exists);
        }

        [TestMethod]
        public void IsFileDeleted()
        {
            dbf.delete();

            bool exists = File.Exists(dbfFileName);
            Assert.IsFalse(exists);
        }
    }
}