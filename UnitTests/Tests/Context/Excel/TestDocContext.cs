using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToDbf.Core.Models;
using ExcelToDbf.Core.Services.Scripts.Data;
using Jint.Native.Function;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests.Tests.Context.Excel
{
    [TestClass]
    public class TestDocContext : AbstractExcel
    {
        [ClassInitialize]
        public static void Init(TestContext testContext) => DefaultInit(testContext);

        protected static DocForm CustomForm(string write, string before=null, string after = null)
        {
            ScriptFunctionInstance MakeFunction(string code, string args="")
            {
                var fullCode = $"(function({args}) {{ \n{code} }})";
                return engine.Evaluate(fullCode) as ScriptFunctionInstance;
            }

            return new DocForm
            {
                Name = "Form",
                BeforeWrite = MakeFunction(before),
                AfterWrite = MakeFunction(after),
                Write = MakeFunction(write, "line"),
            };
        }

        [TestMethod]
        public void TestContextSum()
        {
            var form = CustomForm(@"
                context.sum += parseInt(line[0]) || 0;
                return null;
            ", @"
                context.mustBe = 30;
                context.sum = 0;
            ", @"
                log('Total sum: ' + context.sum);
                if (context.sum != context.mustBe) error('Wrong sum!');
            ");
            var excel = Prepare(form).Connect(null);
            excel.CallHook(form, DocForm.HookType.Before);
            excel.Transform(form, new object[] { "12" });
            excel.Transform(form, new object[] { "18" });
            excel.CallHook(form, DocForm.HookType.After);
        }

    }
}
